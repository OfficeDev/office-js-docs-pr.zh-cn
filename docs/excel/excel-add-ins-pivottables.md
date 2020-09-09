---
title: 使用 Excel JavaScript API 处理数据透视表
description: 使用 Excel JavaScript API 创建数据透视表并与其组件进行交互。
ms.date: 04/20/2020
localization_priority: Normal
ms.openlocfilehash: f1a692f493fc945f114270b69b2ce2004588b731
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2020
ms.locfileid: "47408521"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="e48b6-103">使用 Excel JavaScript API 处理数据透视表</span><span class="sxs-lookup"><span data-stu-id="e48b6-103">Work with PivotTables using the Excel JavaScript API</span></span>

<span data-ttu-id="e48b6-104">数据透视表精简了较大的数据集。</span><span class="sxs-lookup"><span data-stu-id="e48b6-104">PivotTables streamline larger data sets.</span></span> <span data-ttu-id="e48b6-105">它们允许快速操作分组数据。</span><span class="sxs-lookup"><span data-stu-id="e48b6-105">They allow the quick manipulation of grouped data.</span></span> <span data-ttu-id="e48b6-106">Excel JavaScript API 允许你的外接程序创建数据透视表并与其组件进行交互。</span><span class="sxs-lookup"><span data-stu-id="e48b6-106">The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span> <span data-ttu-id="e48b6-107">本文介绍了 Office JavaScript API 如何表示数据透视表，并提供了主要方案的代码示例。</span><span class="sxs-lookup"><span data-stu-id="e48b6-107">This article describes how PivotTables are represented by the Office JavaScript API and provides code samples for key scenarios.</span></span>

<span data-ttu-id="e48b6-108">如果您对数据透视表的功能不熟悉，请考虑将其作为最终用户来浏览。</span><span class="sxs-lookup"><span data-stu-id="e48b6-108">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end user.</span></span>
<span data-ttu-id="e48b6-109">有关这些工具的最佳入门知识，请参阅 [创建数据透视表以分析工作表数据](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) 。</span><span class="sxs-lookup"><span data-stu-id="e48b6-109">See [Create a PivotTable to analyze worksheet data](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="e48b6-110">目前不支持使用 OLAP 创建的数据透视表。</span><span class="sxs-lookup"><span data-stu-id="e48b6-110">PivotTables created with OLAP are not currently supported.</span></span> <span data-ttu-id="e48b6-111">此外，也不支持 Power Pivot。</span><span class="sxs-lookup"><span data-stu-id="e48b6-111">There is also no support for Power Pivot.</span></span>

## <a name="object-model"></a><span data-ttu-id="e48b6-112">对象模型</span><span class="sxs-lookup"><span data-stu-id="e48b6-112">Object model</span></span>

<span data-ttu-id="e48b6-113">[数据透视表](/javascript/api/excel/excel.pivottable)是 OFFICE JavaScript API 中数据透视表的中心对象。</span><span class="sxs-lookup"><span data-stu-id="e48b6-113">The [PivotTable](/javascript/api/excel/excel.pivottable) is the central object for PivotTables in the Office JavaScript API.</span></span>

- <span data-ttu-id="e48b6-114">`Workbook.pivotTables`并且 `Worksheet.pivotTables` 是分别在工作簿和工作表中包含[数据透视](/javascript/api/excel/excel.pivottable)表的[PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) 。</span><span class="sxs-lookup"><span data-stu-id="e48b6-114">`Workbook.pivotTables` and `Worksheet.pivotTables` are [PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) that contain the [PivotTables](/javascript/api/excel/excel.pivottable) in the workbook and worksheet, respectively.</span></span>
- <span data-ttu-id="e48b6-115">[数据透视表](/javascript/api/excel/excel.pivottable)包含具有多个[PivotHierarchies](/javascript/api/excel/excel.pivothierarchy)的[PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection) 。</span><span class="sxs-lookup"><span data-stu-id="e48b6-115">A [PivotTable](/javascript/api/excel/excel.pivottable) contains a [PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection) that has multiple [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy).</span></span>
- <span data-ttu-id="e48b6-116">可以将这些 [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy) 添加到特定的层次结构集合，以定义数据透视表透视数据的方式 (如 [以下部分](#hierarchies) 所述) 。</span><span class="sxs-lookup"><span data-stu-id="e48b6-116">These [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy) can be added to specific hierarchy collections to define how the PivotTable pivots data (as explained in the [following section](#hierarchies)).</span></span>
- <span data-ttu-id="e48b6-117">[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)包含一个仅具有一个[透视字段](/javascript/api/excel/excel.pivotfield)的[PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection) 。</span><span class="sxs-lookup"><span data-stu-id="e48b6-117">A [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) contains a [PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection) that has exactly one [PivotField](/javascript/api/excel/excel.pivotfield).</span></span> <span data-ttu-id="e48b6-118">如果设计扩展以包含 OLAP 数据透视表，则可能会发生更改。</span><span class="sxs-lookup"><span data-stu-id="e48b6-118">If the design expands to include OLAP PivotTables, this may change.</span></span>
- <span data-ttu-id="e48b6-119">[透视字段](/javascript/api/excel/excel.pivotfield)包含具有多个[PivotItems](/javascript/api/excel/excel.pivotitem)的[PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) 。</span><span class="sxs-lookup"><span data-stu-id="e48b6-119">A [PivotField](/javascript/api/excel/excel.pivotfield) contains a [PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) that has multiple [PivotItems](/javascript/api/excel/excel.pivotitem).</span></span>
- <span data-ttu-id="e48b6-120">[数据透视表](/javascript/api/excel/excel.pivottable)包含一个[PivotLayout](/javascript/api/excel/excel.pivotlayout) ，用于定义在工作表中显示[透视字段](/javascript/api/excel/excel.pivotfield)和[PivotItems](/javascript/api/excel/excel.pivotitem)的位置。</span><span class="sxs-lookup"><span data-stu-id="e48b6-120">A [PivotTable](/javascript/api/excel/excel.pivottable) contains a [PivotLayout](/javascript/api/excel/excel.pivotlayout) that defines where the [PivotFields](/javascript/api/excel/excel.pivotfield) and [PivotItems](/javascript/api/excel/excel.pivotitem) are displayed in the worksheet.</span></span>

<span data-ttu-id="e48b6-121">让我们来看看这些关系如何应用于一些示例数据。</span><span class="sxs-lookup"><span data-stu-id="e48b6-121">Let's look at how these relationships apply to some example data.</span></span> <span data-ttu-id="e48b6-122">以下数据介绍了来自不同服务器场的水果销售。</span><span class="sxs-lookup"><span data-stu-id="e48b6-122">The following data describes fruit sales from various farms.</span></span> <span data-ttu-id="e48b6-123">本主题将是本文中的示例。</span><span class="sxs-lookup"><span data-stu-id="e48b6-123">It will be the example throughout this article.</span></span>

![来自不同服务器场的不同类型的水果销售的集合。](../images/excel-pivots-raw-data.png)

<span data-ttu-id="e48b6-125">此水果可使用的销售数据将用于制作数据透视表。</span><span class="sxs-lookup"><span data-stu-id="e48b6-125">This fruit farm sales data will be used to make a PivotTable.</span></span> <span data-ttu-id="e48b6-126">每个列（例如 " **类型**"）是 `PivotHierarchy` 。</span><span class="sxs-lookup"><span data-stu-id="e48b6-126">Each column, such as **Types**, is a `PivotHierarchy`.</span></span> <span data-ttu-id="e48b6-127">" **类型** " 层次结构包含 " **类型** " 字段。</span><span class="sxs-lookup"><span data-stu-id="e48b6-127">The **Types** hierarchy contains the **Types** field.</span></span> <span data-ttu-id="e48b6-128">" **类型** " 字段包含 **苹果**、 **Kiwi**、 **柠檬**、 **酸**橙色和 **橙色**的项。</span><span class="sxs-lookup"><span data-stu-id="e48b6-128">The **Types** field contains the items **Apple**, **Kiwi**, **Lemon**, **Lime**, and **Orange**.</span></span>

### <a name="hierarchies"></a><span data-ttu-id="e48b6-129">Hierarchies</span><span class="sxs-lookup"><span data-stu-id="e48b6-129">Hierarchies</span></span>

<span data-ttu-id="e48b6-130">数据透视表基于四种层次结构类别进行组织： [行](/javascript/api/excel/excel.rowcolumnpivothierarchy)、 [列](/javascript/api/excel/excel.rowcolumnpivothierarchy)、 [数据](/javascript/api/excel/excel.datapivothierarchy)和 [筛选器](/javascript/api/excel/excel.filterpivothierarchy)。</span><span class="sxs-lookup"><span data-stu-id="e48b6-130">PivotTables are organized based on four hierarchy categories: [row](/javascript/api/excel/excel.rowcolumnpivothierarchy), [column](/javascript/api/excel/excel.rowcolumnpivothierarchy), [data](/javascript/api/excel/excel.datapivothierarchy), and [filter](/javascript/api/excel/excel.filterpivothierarchy).</span></span>

<span data-ttu-id="e48b6-131">前面显示的服务器场数据具有五个层次**结构：服务器场、\*\*\*\*类型**、**分类**、**服务器场中销售的 Crates**和**Crates 销售批发**。</span><span class="sxs-lookup"><span data-stu-id="e48b6-131">The farm data shown earlier has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span></span> <span data-ttu-id="e48b6-132">每个层次结构只能存在于四个类别之一中。</span><span class="sxs-lookup"><span data-stu-id="e48b6-132">Each hierarchy can only exist in one of the four categories.</span></span> <span data-ttu-id="e48b6-133">如果 **Type** 添加到列层次结构，则它不能也在行、数据或筛选器层次结构中。</span><span class="sxs-lookup"><span data-stu-id="e48b6-133">If **Type** is added to column hierarchies, it cannot also be in the row, data, or filter hierarchies.</span></span> <span data-ttu-id="e48b6-134">如果 **类型** 随后添加到行层次结构中，则会将其从列层次结构中删除。</span><span class="sxs-lookup"><span data-stu-id="e48b6-134">If **Type** is subsequently added to row hierarchies, it is removed from the column hierarchies.</span></span> <span data-ttu-id="e48b6-135">无论是通过 Excel UI 还是 Excel JavaScript Api 执行层次结构分配，此行为都相同。</span><span class="sxs-lookup"><span data-stu-id="e48b6-135">This behavior is the same whether hierarchy assignment is done through the Excel UI or the Excel JavaScript APIs.</span></span>

<span data-ttu-id="e48b6-136">行和列层次结构定义数据的分组方式。</span><span class="sxs-lookup"><span data-stu-id="e48b6-136">Row and column hierarchies define how data will be grouped.</span></span> <span data-ttu-id="e48b6-137">例如， **服务器场** 的行层次结构将把来自同一个服务器场的所有数据集组合在一起。</span><span class="sxs-lookup"><span data-stu-id="e48b6-137">For example, a row hierarchy of **Farms** will group together all the data sets from the same farm.</span></span> <span data-ttu-id="e48b6-138">在行和列层次结构之间进行选择，以定义数据透视表的方向。</span><span class="sxs-lookup"><span data-stu-id="e48b6-138">The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="e48b6-139">数据层次结构是要根据行和列层次结构聚合的值。</span><span class="sxs-lookup"><span data-stu-id="e48b6-139">Data hierarchies are the values to be aggregated based on the row and column hierarchies.</span></span> <span data-ttu-id="e48b6-140">具有 **服务器场** 的行层次结构和 **Crates** 的数据层次结构的数据透视表显示每个服务器场的所有不同 fruits 的默认) 的总计 (。</span><span class="sxs-lookup"><span data-stu-id="e48b6-140">A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="e48b6-141">筛选器层次结构基于该筛选类型中的值包括或排除数据透视表中的数据。</span><span class="sxs-lookup"><span data-stu-id="e48b6-141">Filter hierarchies include or exclude data from the pivot based on values within that filtered type.</span></span> <span data-ttu-id="e48b6-142">选定类型为 "**有机**" 的**分类**筛选器层次结构仅显示用于随机水果的数据。</span><span class="sxs-lookup"><span data-stu-id="e48b6-142">A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="e48b6-143">下面是数据透视表旁边的服务器场数据。</span><span class="sxs-lookup"><span data-stu-id="e48b6-143">Here is the farm data again, alongside a PivotTable.</span></span> <span data-ttu-id="e48b6-144">数据透视表使用 **服务器场** 和 **类型** 作为行层次结构， **Crates 在服务器场** 和 **Crates** 售出销售批发作为数据 (层次结构，并将其作为) 的默认聚合函数的数据层次结构，并将 **分类** 用作包含 **随机** 选择的) 的筛选器层次结构 (。</span><span class="sxs-lookup"><span data-stu-id="e48b6-144">The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span>

![选择了具有行、数据和筛选器层次结构的数据透视表旁边的水果销售数据。](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="e48b6-146">此数据透视表可通过 JavaScript API 或 Excel UI 生成。</span><span class="sxs-lookup"><span data-stu-id="e48b6-146">This PivotTable could be generated through the JavaScript API or through the Excel UI.</span></span> <span data-ttu-id="e48b6-147">这两个选项都允许通过外接程序进行进一步操作。</span><span class="sxs-lookup"><span data-stu-id="e48b6-147">Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="e48b6-148">创建数据透视表</span><span class="sxs-lookup"><span data-stu-id="e48b6-148">Create a PivotTable</span></span>

<span data-ttu-id="e48b6-149">数据透视表需要名称、源和目标。</span><span class="sxs-lookup"><span data-stu-id="e48b6-149">PivotTables need a name, source, and destination.</span></span> <span data-ttu-id="e48b6-150">源可以是作为 `Range` 、 `string` 或类型) 传递 (的区域地址或表名 `Table` 。</span><span class="sxs-lookup"><span data-stu-id="e48b6-150">The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type).</span></span> <span data-ttu-id="e48b6-151">目标是作为或) 给定的区域地址 (`Range` `string` 。</span><span class="sxs-lookup"><span data-stu-id="e48b6-151">The destination is a range address (given as either a `Range` or `string`).</span></span>
<span data-ttu-id="e48b6-152">下面的示例展示了各种数据透视表创建技术。</span><span class="sxs-lookup"><span data-stu-id="e48b6-152">The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="e48b6-153">创建包含区域地址的数据透视表</span><span class="sxs-lookup"><span data-stu-id="e48b6-153">Create a PivotTable with range addresses</span></span>

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on the current worksheet at cell
    // A22 with data from the range A1:E21.
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add(
      "Farm Sales", "A1:E21", "A22");

    return context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="e48b6-154">创建包含 Range 对象的数据透视表</span><span class="sxs-lookup"><span data-stu-id="e48b6-154">Create a PivotTable with Range objects</span></span>

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

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="e48b6-155">在工作簿级别创建数据透视表</span><span class="sxs-lookup"><span data-stu-id="e48b6-155">Create a PivotTable at the workbook level</span></span>

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21.
    context.workbook.pivotTables.add(
        "Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    return context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="e48b6-156">使用现有数据透视表</span><span class="sxs-lookup"><span data-stu-id="e48b6-156">Use an existing PivotTable</span></span>

<span data-ttu-id="e48b6-157">手动创建的数据透视表也可通过工作簿或单个工作表的数据透视表集合进行访问。</span><span class="sxs-lookup"><span data-stu-id="e48b6-157">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> <span data-ttu-id="e48b6-158">下面的代码从工作簿中获取名为 **"我的透视** 表" 的数据透视表。</span><span class="sxs-lookup"><span data-stu-id="e48b6-158">The following code gets a PivotTable named **My Pivot** from the workbook.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    return context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="e48b6-159">向数据透视表添加行和列</span><span class="sxs-lookup"><span data-stu-id="e48b6-159">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="e48b6-160">数据透视这些字段的值周围的行和列。</span><span class="sxs-lookup"><span data-stu-id="e48b6-160">Rows and columns pivot the data around those fields' values.</span></span>

<span data-ttu-id="e48b6-161">添加 " **服务器场** " 列将每个服务器场的所有销售额枢轴分布。</span><span class="sxs-lookup"><span data-stu-id="e48b6-161">Adding the **Farm** column pivots all the sales around each farm.</span></span> <span data-ttu-id="e48b6-162">添加 " **类型** " 和 " **分类** " 行会根据所售的水果和是否为 "有随机" 来进一步细分数据。</span><span class="sxs-lookup"><span data-stu-id="e48b6-162">Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

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

<span data-ttu-id="e48b6-164">您还可以拥有仅包含行或列的数据透视表。</span><span class="sxs-lookup"><span data-stu-id="e48b6-164">You can also have a PivotTable with only rows or columns.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    return context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="e48b6-165">向数据透视表中添加数据层次结构</span><span class="sxs-lookup"><span data-stu-id="e48b6-165">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="e48b6-166">数据层次结构使用要基于行和列进行组合的信息填充数据透视表。</span><span class="sxs-lookup"><span data-stu-id="e48b6-166">Data hierarchies fill the PivotTable with information to combine based on the rows and columns.</span></span> <span data-ttu-id="e48b6-167">添加 Crates 的数据层次结构 **在服务器场** 和 **Crates** 售出销售批发为每个行和列提供这些数字的总和。</span><span class="sxs-lookup"><span data-stu-id="e48b6-167">Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span>

<span data-ttu-id="e48b6-168">在示例中，" **服务器场** " 和 " **类型** " 都是行，而 "发货箱销售额" 作为数据。</span><span class="sxs-lookup"><span data-stu-id="e48b6-168">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span>

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

## <a name="pivottable-layouts-and-getting-pivoted-data"></a><span data-ttu-id="e48b6-170">数据透视表布局和获取透视数据</span><span class="sxs-lookup"><span data-stu-id="e48b6-170">PivotTable layouts and getting pivoted data</span></span>

<span data-ttu-id="e48b6-171">[PivotLayout](/javascript/api/excel/excel.pivotlayout)定义层次结构及其数据的位置。</span><span class="sxs-lookup"><span data-stu-id="e48b6-171">A [PivotLayout](/javascript/api/excel/excel.pivotlayout) defines the placement of hierarchies and their data.</span></span> <span data-ttu-id="e48b6-172">您可以访问布局以确定存储数据的区域。</span><span class="sxs-lookup"><span data-stu-id="e48b6-172">You access the layout to determine the ranges where data is stored.</span></span>

<span data-ttu-id="e48b6-173">下图显示了哪些布局函数调用对应于数据透视表的区域。</span><span class="sxs-lookup"><span data-stu-id="e48b6-173">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![显示由布局的 get range 函数返回的数据透视表的节的图表。](../images/excel-pivots-layout-breakdown.png)

### <a name="get-data-from-the-pivottable"></a><span data-ttu-id="e48b6-175">从数据透视表中获取数据</span><span class="sxs-lookup"><span data-stu-id="e48b6-175">Get data from the PivotTable</span></span>

<span data-ttu-id="e48b6-176">布局定义了数据透视表在工作表中的显示方式。</span><span class="sxs-lookup"><span data-stu-id="e48b6-176">The layout defines how the PivotTable is displayed in the worksheet.</span></span> <span data-ttu-id="e48b6-177">这意味着 `PivotLayout` 对象控制用于数据透视表元素的区域。</span><span class="sxs-lookup"><span data-stu-id="e48b6-177">This means the `PivotLayout` object controls the ranges used for PivotTable elements.</span></span> <span data-ttu-id="e48b6-178">使用由布局提供的区域来获取由数据透视表收集和聚合的数据。</span><span class="sxs-lookup"><span data-stu-id="e48b6-178">Use the ranges provided by the layout to get data collected and aggregated by the PivotTable.</span></span> <span data-ttu-id="e48b6-179">尤其是，使用 `PivotLayout.getDataBodyRange` 可访问数据透视表所生成的内容。</span><span class="sxs-lookup"><span data-stu-id="e48b6-179">In particular, use `PivotLayout.getDataBodyRange` to access what the PivotTable produces.</span></span>

<span data-ttu-id="e48b6-180">下面的代码演示了如何通过布局来获取数据透视表数据的最后一行， (在**服务器场中售出的 Crates 总和**和) 的早期示例中的**Crates 售出的批发**列的**总和总计。**</span><span class="sxs-lookup"><span data-stu-id="e48b6-180">The following code demonstrates how to get the last row of the PivotTable data by going through the layout (the **Grand Total** of both the **Sum of Crates Sold at Farm** and **Sum of Crates Sold Wholesale** columns in the earlier example).</span></span> <span data-ttu-id="e48b6-181">然后，将这些值汇总到一起，以得到最终总计，在单元格 **E30** 中显示在数据透视表) 外部 (。</span><span class="sxs-lookup"><span data-stu-id="e48b6-181">Those values are then summed together for a final total, which is displayed in cell **E30** (outside of the PivotTable).</span></span>

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

### <a name="layout-types"></a><span data-ttu-id="e48b6-182">布局类型</span><span class="sxs-lookup"><span data-stu-id="e48b6-182">Layout types</span></span>

<span data-ttu-id="e48b6-183">数据透视表具有三种布局样式：紧凑、大纲和表格。</span><span class="sxs-lookup"><span data-stu-id="e48b6-183">PivotTables have three layout styles: Compact, Outline, and Tabular.</span></span> <span data-ttu-id="e48b6-184">我们在前面的示例中看到了压缩样式。</span><span class="sxs-lookup"><span data-stu-id="e48b6-184">We've seen the compact style in the previous examples.</span></span>

<span data-ttu-id="e48b6-185">下面的示例分别使用大纲样式和表格样式。</span><span class="sxs-lookup"><span data-stu-id="e48b6-185">The following examples use the outline and tabular styles, respectively.</span></span> <span data-ttu-id="e48b6-186">此代码示例演示如何在不同的布局之间循环。</span><span class="sxs-lookup"><span data-stu-id="e48b6-186">The code sample shows how to cycle between the different layouts.</span></span>

#### <a name="outline-layout"></a><span data-ttu-id="e48b6-187">大纲布局</span><span class="sxs-lookup"><span data-stu-id="e48b6-187">Outline layout</span></span>

![使用大纲布局的数据透视表。](../images/excel-pivots-outline-layout.png)

#### <a name="tabular-layout"></a><span data-ttu-id="e48b6-189">表格布局</span><span class="sxs-lookup"><span data-stu-id="e48b6-189">Tabular layout</span></span>

![使用表格布局的数据透视表。](../images/excel-pivots-tabular-layout.png)

## <a name="delete-a-pivottable"></a><span data-ttu-id="e48b6-191">删除数据透视表</span><span class="sxs-lookup"><span data-stu-id="e48b6-191">Delete a PivotTable</span></span>

<span data-ttu-id="e48b6-192">使用它们的名称删除数据透视表。</span><span class="sxs-lookup"><span data-stu-id="e48b6-192">PivotTables are deleted by using their name.</span></span>

```js
Excel.run(function (context) {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();
    return context.sync();
});
```

## <a name="slicers"></a><span data-ttu-id="e48b6-193">切片器</span><span class="sxs-lookup"><span data-stu-id="e48b6-193">Slicers</span></span>

<span data-ttu-id="e48b6-194">[切片](/javascript/api/excel/excel.slicer) 器允许从 Excel 数据透视表或表中筛选数据。</span><span class="sxs-lookup"><span data-stu-id="e48b6-194">[Slicers](/javascript/api/excel/excel.slicer) allow data to be filtered from an Excel PivotTable or table.</span></span> <span data-ttu-id="e48b6-195">切片器使用指定的列或透视字段中的值筛选相应的行。</span><span class="sxs-lookup"><span data-stu-id="e48b6-195">A slicer uses values from a specified column or PivotField to filter corresponding rows.</span></span> <span data-ttu-id="e48b6-196">这些值存储为中的 [SlicerItem](/javascript/api/excel/excel.sliceritem) 对象 `Slicer` 。</span><span class="sxs-lookup"><span data-stu-id="e48b6-196">These values are stored as [SlicerItem](/javascript/api/excel/excel.sliceritem) objects in the `Slicer`.</span></span> <span data-ttu-id="e48b6-197">你的外接程序可以调整这些筛选器，因为用户可以 [通过 EXCEL UI](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d))  (。</span><span class="sxs-lookup"><span data-stu-id="e48b6-197">Your add-in can adjust these filters, as can users ([through the Excel UI](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)).</span></span> <span data-ttu-id="e48b6-198">切片器位于绘图层中的工作表的顶部，如以下屏幕截图所示。</span><span class="sxs-lookup"><span data-stu-id="e48b6-198">The slicer sits on top of the worksheet in the drawing layer, as shown in the following screenshot.</span></span>

![数据透视表上的切片器筛选数据。](../images/excel-slicer.png)

> [!NOTE]
> <span data-ttu-id="e48b6-200">本节中介绍的技术重点介绍如何使用连接到数据透视表的切片器。</span><span class="sxs-lookup"><span data-stu-id="e48b6-200">The techniques described in this section focus on how to use slicers connected to PivotTables.</span></span> <span data-ttu-id="e48b6-201">同样的技术也适用于使用连接到表的切片器。</span><span class="sxs-lookup"><span data-stu-id="e48b6-201">The same techniques also apply to using slicers connected to tables.</span></span>

### <a name="create-a-slicer"></a><span data-ttu-id="e48b6-202">创建切片器</span><span class="sxs-lookup"><span data-stu-id="e48b6-202">Create a slicer</span></span>

<span data-ttu-id="e48b6-203">您可以使用方法或方法在工作簿或工作表中创建切片器 `Workbook.slicers.add` `Worksheet.slicers.add` 。</span><span class="sxs-lookup"><span data-stu-id="e48b6-203">You can create a slicer in a workbook or worksheet by using the `Workbook.slicers.add` method or `Worksheet.slicers.add` method.</span></span> <span data-ttu-id="e48b6-204">这样做会将切片器添加[SlicerCollection](/javascript/api/excel/excel.slicercollection)到指定 `Workbook` 对象或对象的 SlicerCollection `Worksheet` 。</span><span class="sxs-lookup"><span data-stu-id="e48b6-204">Doing so adds a slicer to the [SlicerCollection](/javascript/api/excel/excel.slicercollection) of the specified `Workbook` or `Worksheet` object.</span></span> <span data-ttu-id="e48b6-205">该 `SlicerCollection.add` 方法具有三个参数：</span><span class="sxs-lookup"><span data-stu-id="e48b6-205">The `SlicerCollection.add` method has three parameters:</span></span>

- <span data-ttu-id="e48b6-206">`slicerSource`：新切片器所基于的数据源。</span><span class="sxs-lookup"><span data-stu-id="e48b6-206">`slicerSource`: The data source on which the new slicer is based.</span></span> <span data-ttu-id="e48b6-207">它可以是 `PivotTable` 、或 `Table` 字符串，代表或的名称或 ID `PivotTable` `Table` 。</span><span class="sxs-lookup"><span data-stu-id="e48b6-207">It can be a `PivotTable`, `Table`, or string representing the name or ID of a `PivotTable` or `Table`.</span></span>
- <span data-ttu-id="e48b6-208">`sourceField`：要作为筛选依据的数据源中的字段。</span><span class="sxs-lookup"><span data-stu-id="e48b6-208">`sourceField`: The field in the data source by which to filter.</span></span> <span data-ttu-id="e48b6-209">它可以是 `PivotField` 、或 `TableColumn` 字符串，代表或的名称或 ID `PivotField` `TableColumn` 。</span><span class="sxs-lookup"><span data-stu-id="e48b6-209">It can be a `PivotField`, `TableColumn`, or string representing the name or ID of a `PivotField` or `TableColumn`.</span></span>
- <span data-ttu-id="e48b6-210">`slicerDestination`：将在其中创建新切片器的工作表。</span><span class="sxs-lookup"><span data-stu-id="e48b6-210">`slicerDestination`: The worksheet where the new slicer will be created.</span></span> <span data-ttu-id="e48b6-211">它可以是一个对象，也可以是的 `Worksheet` 名称或 ID `Worksheet` 。</span><span class="sxs-lookup"><span data-stu-id="e48b6-211">It can be a `Worksheet` object or the name or ID of a `Worksheet`.</span></span> <span data-ttu-id="e48b6-212">通过访问时，此参数是不必要的 `SlicerCollection` `Worksheet.slicers` 。</span><span class="sxs-lookup"><span data-stu-id="e48b6-212">This parameter is unnecessary when the `SlicerCollection` is accessed through `Worksheet.slicers`.</span></span> <span data-ttu-id="e48b6-213">在这种情况下，集合的工作表将用作目标。</span><span class="sxs-lookup"><span data-stu-id="e48b6-213">In this case, the collection's worksheet is used as the destination.</span></span>

<span data-ttu-id="e48b6-214">下面的代码示例向 **数据透视** 表中添加一个新的切片器。</span><span class="sxs-lookup"><span data-stu-id="e48b6-214">The following code sample adds a new slicer to the **Pivot** worksheet.</span></span> <span data-ttu-id="e48b6-215">切片器的源是 **服务器场销售** 数据透视表和使用 **类型** 数据的筛选器。</span><span class="sxs-lookup"><span data-stu-id="e48b6-215">The slicer's source is the **Farm Sales** PivotTable and filters using the **Type** data.</span></span> <span data-ttu-id="e48b6-216">切片器也称为 **水果切片器** ，以供将来参考。</span><span class="sxs-lookup"><span data-stu-id="e48b6-216">The slicer is also named **Fruit Slicer** for future reference.</span></span>

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

### <a name="filter-items-with-a-slicer"></a><span data-ttu-id="e48b6-217">使用切片器筛选项目</span><span class="sxs-lookup"><span data-stu-id="e48b6-217">Filter items with a slicer</span></span>

<span data-ttu-id="e48b6-218">切片器使用中的项筛选数据透视表 `sourceField` 。</span><span class="sxs-lookup"><span data-stu-id="e48b6-218">The slicer filters the PivotTable with items from the `sourceField`.</span></span> <span data-ttu-id="e48b6-219">该 `Slicer.selectItems` 方法将设置切片器中保留的项。</span><span class="sxs-lookup"><span data-stu-id="e48b6-219">The `Slicer.selectItems` method sets the items that remain in the slicer.</span></span> <span data-ttu-id="e48b6-220">这些项作为 a 传递给方法 `string[]` ，表示项的键。</span><span class="sxs-lookup"><span data-stu-id="e48b6-220">These items are passed to the method as a `string[]`, representing the keys of the items.</span></span> <span data-ttu-id="e48b6-221">包含这些项目的任何行仍保留在数据透视表的聚合中。</span><span class="sxs-lookup"><span data-stu-id="e48b6-221">Any rows containing those items remain in the PivotTable's aggregation.</span></span> <span data-ttu-id="e48b6-222">随后调用 `selectItems` 将列表设置为在这些调用中指定的键。</span><span class="sxs-lookup"><span data-stu-id="e48b6-222">Subsequent calls to `selectItems` set the list to the keys specified in those calls.</span></span>

> [!NOTE]
> <span data-ttu-id="e48b6-223">如果 `Slicer.selectItems` 向传递的项不在数据源中，则 `InvalidArgument` 会引发错误。</span><span class="sxs-lookup"><span data-stu-id="e48b6-223">If `Slicer.selectItems` is passed an item that's not in the data source, an `InvalidArgument` error is thrown.</span></span> <span data-ttu-id="e48b6-224">可以通过属性来验证内容 `Slicer.slicerItems` ，这是 [SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection)。</span><span class="sxs-lookup"><span data-stu-id="e48b6-224">The contents can be verified through the `Slicer.slicerItems` property, which is a [SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection).</span></span>

<span data-ttu-id="e48b6-225">下面的代码示例显示为切片器选择了三个项目： **柠檬**、 **酸橙色**和 **橙色**。</span><span class="sxs-lookup"><span data-stu-id="e48b6-225">The following code sample shows three items being selected for the slicer: **Lemon**, **Lime**, and **Orange**.</span></span>

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    // Anything other than the following three values will be filtered out of the PivotTable for display and aggregation.
    slicer.selectItems(["Lemon", "Lime", "Orange"]);
    return context.sync();
});
```

<span data-ttu-id="e48b6-226">若要从切片器中删除所有筛选器，请使用 `Slicer.clearFilters` 方法，如下面的示例所示。</span><span class="sxs-lookup"><span data-stu-id="e48b6-226">To remove all filters from the slicer, use the `Slicer.clearFilters` method, as shown in the following sample.</span></span>

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.clearFilters();
    return context.sync();
});
```

### <a name="style-and-format-a-slicer"></a><span data-ttu-id="e48b6-227">为切片器设置样式和格式</span><span class="sxs-lookup"><span data-stu-id="e48b6-227">Style and format a slicer</span></span>

<span data-ttu-id="e48b6-228">您的外接可以通过属性调整切片器的显示设置 `Slicer` 。</span><span class="sxs-lookup"><span data-stu-id="e48b6-228">You add-in can adjust a slicer's display settings through `Slicer` properties.</span></span> <span data-ttu-id="e48b6-229">下面的代码示例将样式设置为 **SlicerStyleLight6**，将切片器顶部的文本设置为 **水果类型**，将切片器放置在绘图层上 \*\* (395，15) \*\* 的位置，并将切片器的大小设置为 **135x150** 像素。</span><span class="sxs-lookup"><span data-stu-id="e48b6-229">The following code sample sets the style to **SlicerStyleLight6**, sets the text at the top of the slicer to **Fruit Types**, places the slicer at the position **(395, 15)** on the drawing layer, and sets the slicer's size to **135x150** pixels.</span></span>

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

### <a name="delete-a-slicer"></a><span data-ttu-id="e48b6-230">删除切片器</span><span class="sxs-lookup"><span data-stu-id="e48b6-230">Delete a slicer</span></span>

<span data-ttu-id="e48b6-231">若要删除切片器，请调用 `Slicer.delete` 方法。</span><span class="sxs-lookup"><span data-stu-id="e48b6-231">To delete a slicer, call the `Slicer.delete` method.</span></span> <span data-ttu-id="e48b6-232">下面的代码示例从当前工作表中删除第一个切片器。</span><span class="sxs-lookup"><span data-stu-id="e48b6-232">The following code sample deletes the first slicer from the current worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.slicers.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="change-aggregation-function"></a><span data-ttu-id="e48b6-233">更改聚合函数</span><span class="sxs-lookup"><span data-stu-id="e48b6-233">Change aggregation function</span></span>

<span data-ttu-id="e48b6-234">数据层次结构的值已聚合。</span><span class="sxs-lookup"><span data-stu-id="e48b6-234">Data hierarchies have their values aggregated.</span></span> <span data-ttu-id="e48b6-235">对于数字的数据集，默认情况下，这是一个总和。</span><span class="sxs-lookup"><span data-stu-id="e48b6-235">For datasets of numbers, this is a sum by default.</span></span> <span data-ttu-id="e48b6-236">该 `summarizeBy` 属性基于 [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) 类型定义此行为。</span><span class="sxs-lookup"><span data-stu-id="e48b6-236">The `summarizeBy` property defines this behavior based on an [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) type.</span></span>

<span data-ttu-id="e48b6-237">当前受支持的聚合函数类型分别为、、、、、、、、、、 `Sum` `Count` `Average` `Max` `Min` `Product` `CountNumbers` `StandardDeviation` `StandardDeviationP` `Variance` `VarianceP` 和 `Automatic` (默认的) 。</span><span class="sxs-lookup"><span data-stu-id="e48b6-237">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="e48b6-238">下面的代码示例将聚合更改为数据的平均值。</span><span class="sxs-lookup"><span data-stu-id="e48b6-238">The following code samples changes the aggregation to be averages of the data.</span></span>

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

## <a name="change-calculations-with-a-showasrule"></a><span data-ttu-id="e48b6-239">使用 ShowAsRule 更改计算</span><span class="sxs-lookup"><span data-stu-id="e48b6-239">Change calculations with a ShowAsRule</span></span>

<span data-ttu-id="e48b6-240">默认情况下，数据透视表将单独聚合其行和列层次结构的数据。</span><span class="sxs-lookup"><span data-stu-id="e48b6-240">PivotTables, by default, aggregate the data of their row and column hierarchies independently.</span></span> <span data-ttu-id="e48b6-241">[ShowAsRule](/javascript/api/excel/excel.showasrule)将数据层次结构更改为基于数据透视表中的其他项的输出值。</span><span class="sxs-lookup"><span data-stu-id="e48b6-241">A [ShowAsRule](/javascript/api/excel/excel.showasrule) changes the data hierarchy to output values based on other items in the PivotTable.</span></span>

<span data-ttu-id="e48b6-242">`ShowAsRule`对象具有三个属性：</span><span class="sxs-lookup"><span data-stu-id="e48b6-242">The `ShowAsRule` object has three properties:</span></span>

- <span data-ttu-id="e48b6-243">`calculation`：要应用于数据层次结构 (的相对计算的类型，默认值为 `none`) 。</span><span class="sxs-lookup"><span data-stu-id="e48b6-243">`calculation`: The type of relative calculation to apply to the data hierarchy (the default is `none`).</span></span>
- <span data-ttu-id="e48b6-244">`baseField`：在应用计算之前包含基础数据的层次结构中的 [透视字段](/javascript/api/excel/excel.pivotfield) 。</span><span class="sxs-lookup"><span data-stu-id="e48b6-244">`baseField`: The [PivotField](/javascript/api/excel/excel.pivotfield) in the hierarchy containing the base data before the calculation is applied.</span></span> <span data-ttu-id="e48b6-245">由于 Excel 数据透视表的层次结构与字段的一对一映射，因此将使用相同的名称来访问层次结构和字段。</span><span class="sxs-lookup"><span data-stu-id="e48b6-245">Since Excel PivotTables have a one-to-one mapping of hierarchy to field, you'll use the same name to access both the hierarchy and the field.</span></span>
- <span data-ttu-id="e48b6-246">`baseItem`：个人 [PivotItem](/javascript/api/excel/excel.pivotitem) 根据计算类型与基本字段的值进行比较。</span><span class="sxs-lookup"><span data-stu-id="e48b6-246">`baseItem`: The individual [PivotItem](/javascript/api/excel/excel.pivotitem) compared against the values of the base fields based on the calculation type.</span></span> <span data-ttu-id="e48b6-247">并非所有计算都需要此字段。</span><span class="sxs-lookup"><span data-stu-id="e48b6-247">Not all calculations require this field.</span></span>

<span data-ttu-id="e48b6-248">以下示例将场数据层次结构中的 " **Crates** 总数" 的计算设置为列总计的百分比。</span><span class="sxs-lookup"><span data-stu-id="e48b6-248">The following example sets the calculation on the **Sum of Crates Sold at Farm** data hierarchy to be a percentage of the column total.</span></span>
<span data-ttu-id="e48b6-249">我们仍希望将粒度扩展到水果类型级别，因此我们将使用 **类型** 行层次结构及其基础字段。</span><span class="sxs-lookup"><span data-stu-id="e48b6-249">We still want the granularity to extend to the fruit type level, so we'll use the **Type** row hierarchy and its underlying field.</span></span>
<span data-ttu-id="e48b6-250">该示例还将 **服务器场** 作为第一个行的层次结构，因此服务器场总数将显示每个服务器场也负责生成的百分比。</span><span class="sxs-lookup"><span data-stu-id="e48b6-250">The example also has **Farm** as the first row hierarchy, so the farm total entries display the percentage each farm is responsible for producing as well.</span></span>

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

<span data-ttu-id="e48b6-252">上面的示例将计算设置为相对于单个行层次结构的字段的列。</span><span class="sxs-lookup"><span data-stu-id="e48b6-252">The previous example set the calculation to the column, relative to the field of an individual row hierarchy.</span></span> <span data-ttu-id="e48b6-253">当计算与单个项目相关时，请使用 `baseItem` 属性。</span><span class="sxs-lookup"><span data-stu-id="e48b6-253">When the calculation relates to an individual item, use the `baseItem` property.</span></span>

<span data-ttu-id="e48b6-254">下面的示例演示了 `differenceFrom` 计算。</span><span class="sxs-lookup"><span data-stu-id="e48b6-254">The following example shows the `differenceFrom` calculation.</span></span> <span data-ttu-id="e48b6-255">它显示场中相对于 **服务器场**中的销售数据层次结构条目的不同之处。</span><span class="sxs-lookup"><span data-stu-id="e48b6-255">It displays the difference of the farm crate sales data hierarchy entries relative to those of **A Farms**.</span></span>
<span data-ttu-id="e48b6-256">`baseField`是**服务器场**，因此我们看到其他服务器场之间的差异，以及每种类型的类似水果的细目 (**类型**也是此示例中的行层次结构) 。</span><span class="sxs-lookup"><span data-stu-id="e48b6-256">The `baseField` is **Farm**, so we see the differences between the other farms, as well as breakdowns for each type of like fruit (**Type** is also a row hierarchy in this example).</span></span>

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

## <a name="change-hierarchy-names"></a><span data-ttu-id="e48b6-260">更改层次结构名称</span><span class="sxs-lookup"><span data-stu-id="e48b6-260">Change hierarchy names</span></span>

<span data-ttu-id="e48b6-261">层次结构字段是可编辑的。</span><span class="sxs-lookup"><span data-stu-id="e48b6-261">Hierarchy fields are editable.</span></span> <span data-ttu-id="e48b6-262">下面的代码演示如何更改两个数据层次结构的显示名称。</span><span class="sxs-lookup"><span data-stu-id="e48b6-262">The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="e48b6-263">另请参阅</span><span class="sxs-lookup"><span data-stu-id="e48b6-263">See also</span></span>

- [<span data-ttu-id="e48b6-264">Office 外接程序中的 Excel JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="e48b6-264">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="e48b6-265">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="e48b6-265">Excel JavaScript API Reference</span></span>](/javascript/api/excel)
