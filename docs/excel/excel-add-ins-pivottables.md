---
title: 使用 JavaScript API Excel数据透视表
description: 使用 Excel JavaScript API 创建数据透视表并与其组件交互。
ms.date: 07/02/2021
localization_priority: Normal
ms.openlocfilehash: 8c8917f57b7546694e12380fc4369847be24ceac
ms.sourcegitcommit: aa73ec6367eaf74399fbf8d6b7776d77895e9982
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/03/2021
ms.locfileid: "53290738"
---
# <a name="work-with-pivottables-using-the-excel-javascript-api"></a><span data-ttu-id="350d9-103">使用 JavaScript API Excel数据透视表</span><span class="sxs-lookup"><span data-stu-id="350d9-103">Work with PivotTables using the Excel JavaScript API</span></span>

<span data-ttu-id="350d9-104">数据透视表可简化更大的数据集。</span><span class="sxs-lookup"><span data-stu-id="350d9-104">PivotTables streamline larger data sets.</span></span> <span data-ttu-id="350d9-105">它们允许对分组数据进行快速操作。</span><span class="sxs-lookup"><span data-stu-id="350d9-105">They allow the quick manipulation of grouped data.</span></span> <span data-ttu-id="350d9-106">借助 Excel JavaScript API，加载项可以创建数据透视表并与其组件交互。</span><span class="sxs-lookup"><span data-stu-id="350d9-106">The Excel JavaScript API lets your add-in create PivotTables and interact with their components.</span></span> <span data-ttu-id="350d9-107">本文介绍数据透视表如何由 Office JavaScript API 表示，并提供关键方案的代码示例。</span><span class="sxs-lookup"><span data-stu-id="350d9-107">This article describes how PivotTables are represented by the Office JavaScript API and provides code samples for key scenarios.</span></span>

<span data-ttu-id="350d9-108">如果您不熟悉数据透视表的功能，请考虑以最终用户模式探索它们。</span><span class="sxs-lookup"><span data-stu-id="350d9-108">If you are unfamiliar with the functionality of PivotTables, consider exploring them as an end user.</span></span>
<span data-ttu-id="350d9-109">有关 [这些工具的良好基础，](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) 请参阅创建数据透视表以分析工作表数据。</span><span class="sxs-lookup"><span data-stu-id="350d9-109">See [Create a PivotTable to analyze worksheet data](https://support.office.com/article/Import-and-analyze-data-ccd3c4a6-272f-4c97-afbb-d3f27407fcde#ID0EAABAAA=PivotTables) for a good primer on these tools.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="350d9-110">当前不支持使用 OLAP 创建的数据透视表。</span><span class="sxs-lookup"><span data-stu-id="350d9-110">PivotTables created with OLAP are not currently supported.</span></span> <span data-ttu-id="350d9-111">也不支持 Power Pivot。</span><span class="sxs-lookup"><span data-stu-id="350d9-111">There is also no support for Power Pivot.</span></span>

## <a name="object-model"></a><span data-ttu-id="350d9-112">对象模型</span><span class="sxs-lookup"><span data-stu-id="350d9-112">Object model</span></span>

<span data-ttu-id="350d9-113">数据[透视表](/javascript/api/excel/excel.pivottable)是 JavaScript API 中数据透视表Office对象。</span><span class="sxs-lookup"><span data-stu-id="350d9-113">The [PivotTable](/javascript/api/excel/excel.pivottable) is the central object for PivotTables in the Office JavaScript API.</span></span>

- <span data-ttu-id="350d9-114">`Workbook.pivotTables`和 是分别包含工作簿和工作表中的 `Worksheet.pivotTables` 数据透视[](/javascript/api/excel/excel.pivottable)表的[PivotTableCollection。](/javascript/api/excel/excel.pivottablecollection)</span><span class="sxs-lookup"><span data-stu-id="350d9-114">`Workbook.pivotTables` and `Worksheet.pivotTables` are [PivotTableCollections](/javascript/api/excel/excel.pivottablecollection) that contain the [PivotTables](/javascript/api/excel/excel.pivottable) in the workbook and worksheet, respectively.</span></span>
- <span data-ttu-id="350d9-115">数据[透视表](/javascript/api/excel/excel.pivottable)包含具有[多个 PivotHierarchies 的 PivotHierarchyCollection。](/javascript/api/excel/excel.pivothierarchycollection) [](/javascript/api/excel/excel.pivothierarchy)</span><span class="sxs-lookup"><span data-stu-id="350d9-115">A [PivotTable](/javascript/api/excel/excel.pivottable) contains a [PivotHierarchyCollection](/javascript/api/excel/excel.pivothierarchycollection) that has multiple [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy).</span></span>
- <span data-ttu-id="350d9-116">可以将[这些 PivotHierarchies](/javascript/api/excel/excel.pivothierarchy)添加到特定层次结构集合中，以定义数据透视表 (数据透视表数据透视表) 。 [](#hierarchies)</span><span class="sxs-lookup"><span data-stu-id="350d9-116">These [PivotHierarchies](/javascript/api/excel/excel.pivothierarchy) can be added to specific hierarchy collections to define how the PivotTable pivots data (as explained in the [following section](#hierarchies)).</span></span>
- <span data-ttu-id="350d9-117">[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)包含一个只有一个 PivotField 的[PivotFieldCollection。](/javascript/api/excel/excel.pivotfieldcollection) [](/javascript/api/excel/excel.pivotfield)</span><span class="sxs-lookup"><span data-stu-id="350d9-117">A [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) contains a [PivotFieldCollection](/javascript/api/excel/excel.pivotfieldcollection) that has exactly one [PivotField](/javascript/api/excel/excel.pivotfield).</span></span> <span data-ttu-id="350d9-118">如果设计扩展为包含 OLAP 数据透视表，这可能会更改。</span><span class="sxs-lookup"><span data-stu-id="350d9-118">If the design expands to include OLAP PivotTables, this may change.</span></span>
- <span data-ttu-id="350d9-119">只要[将字段](/javascript/api/excel/excel.pivotfield)的[PivotHierarchy](/javascript/api/excel/excel.pivothierarchy)分配给层次结构类别，透视字段就可以应用一个或多个[PivotFilter。](/javascript/api/excel/excel.pivotfilters)</span><span class="sxs-lookup"><span data-stu-id="350d9-119">A [PivotField](/javascript/api/excel/excel.pivotfield) can have one or more [PivotFilters](/javascript/api/excel/excel.pivotfilters) applied, as long as the field's [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) is assigned to a hierarchy category.</span></span>
- <span data-ttu-id="350d9-120">透视[字段](/javascript/api/excel/excel.pivotfield)包含具有[多个 PivotItems 的 PivotItemCollection。](/javascript/api/excel/excel.pivotitemcollection) [](/javascript/api/excel/excel.pivotitem)</span><span class="sxs-lookup"><span data-stu-id="350d9-120">A [PivotField](/javascript/api/excel/excel.pivotfield) contains a [PivotItemCollection](/javascript/api/excel/excel.pivotitemcollection) that has multiple [PivotItems](/javascript/api/excel/excel.pivotitem).</span></span>
- <span data-ttu-id="350d9-121">[数据透视表](/javascript/api/excel/excel.pivottable)包含[一个 PivotLayout，](/javascript/api/excel/excel.pivotlayout)它定义[透视字段](/javascript/api/excel/excel.pivotfield)和[PivotItems](/javascript/api/excel/excel.pivotitem)在工作表中的显示位置。</span><span class="sxs-lookup"><span data-stu-id="350d9-121">A [PivotTable](/javascript/api/excel/excel.pivottable) contains a [PivotLayout](/javascript/api/excel/excel.pivotlayout) that defines where the [PivotFields](/javascript/api/excel/excel.pivotfield) and [PivotItems](/javascript/api/excel/excel.pivotitem) are displayed in the worksheet.</span></span> <span data-ttu-id="350d9-122">布局还控制数据透视表的一些显示设置。</span><span class="sxs-lookup"><span data-stu-id="350d9-122">The layout also controls some display settings for the PivotTable.</span></span>

<span data-ttu-id="350d9-123">让我们看一下这些关系如何应用于一些示例数据。</span><span class="sxs-lookup"><span data-stu-id="350d9-123">Let's look at how these relationships apply to some example data.</span></span> <span data-ttu-id="350d9-124">以下数据描述了来自各种服务器场的菜品销售情况。</span><span class="sxs-lookup"><span data-stu-id="350d9-124">The following data describes fruit sales from various farms.</span></span> <span data-ttu-id="350d9-125">它将是整篇文章中的示例。</span><span class="sxs-lookup"><span data-stu-id="350d9-125">It will be the example throughout this article.</span></span>

![不同服务器场中不同类型的新鲜品销售的集合。](../images/excel-pivots-raw-data.png)

<span data-ttu-id="350d9-127">此菜场销售数据将用于创建数据透视表。</span><span class="sxs-lookup"><span data-stu-id="350d9-127">This fruit farm sales data will be used to make a PivotTable.</span></span> <span data-ttu-id="350d9-128">每列（如 **Types）** 都是 `PivotHierarchy` 。</span><span class="sxs-lookup"><span data-stu-id="350d9-128">Each column, such as **Types**, is a `PivotHierarchy`.</span></span> <span data-ttu-id="350d9-129">" **类型** "层次结构包含 **"类型"** 字段。</span><span class="sxs-lookup"><span data-stu-id="350d9-129">The **Types** hierarchy contains the **Types** field.</span></span> <span data-ttu-id="350d9-130">The **Types** field contains the items **Apple**， **Kiwi**， **Orange**， **Orange**， and **Orange**.</span><span class="sxs-lookup"><span data-stu-id="350d9-130">The **Types** field contains the items **Apple**, **Kiwi**, **Lemon**, **Lime**, and **Orange**.</span></span>

### <a name="hierarchies"></a><span data-ttu-id="350d9-131">Hierarchies</span><span class="sxs-lookup"><span data-stu-id="350d9-131">Hierarchies</span></span>

<span data-ttu-id="350d9-132">数据透视表基于四个层次结构类别进行组织：[行](/javascript/api/excel/excel.rowcolumnpivothierarchy)、[列](/javascript/api/excel/excel.rowcolumnpivothierarchy)[、数据和](/javascript/api/excel/excel.datapivothierarchy)[筛选器](/javascript/api/excel/excel.filterpivothierarchy)。</span><span class="sxs-lookup"><span data-stu-id="350d9-132">PivotTables are organized based on four hierarchy categories: [row](/javascript/api/excel/excel.rowcolumnpivothierarchy), [column](/javascript/api/excel/excel.rowcolumnpivothierarchy), [data](/javascript/api/excel/excel.datapivothierarchy), and [filter](/javascript/api/excel/excel.filterpivothierarchy).</span></span>

<span data-ttu-id="350d9-133">前面显示的服务器场数据有五个层次结构：Farms、Type、Classification、Crates **Sold at Farm** 和 **Crates SoldRate。**   </span><span class="sxs-lookup"><span data-stu-id="350d9-133">The farm data shown earlier has five hierarchies: **Farms**, **Type**, **Classification**, **Crates Sold at Farm**, and **Crates Sold Wholesale**.</span></span> <span data-ttu-id="350d9-134">每个层次结构只能存在于四个类别之一中。</span><span class="sxs-lookup"><span data-stu-id="350d9-134">Each hierarchy can only exist in one of the four categories.</span></span> <span data-ttu-id="350d9-135">如果将 **Type** 添加到列层次结构中，则不能也位于行、数据或筛选器层次结构中。</span><span class="sxs-lookup"><span data-stu-id="350d9-135">If **Type** is added to column hierarchies, it cannot also be in the row, data, or filter hierarchies.</span></span> <span data-ttu-id="350d9-136">如果 **Type** 随后添加到行层次结构中，则从列层次结构中删除它。</span><span class="sxs-lookup"><span data-stu-id="350d9-136">If **Type** is subsequently added to row hierarchies, it is removed from the column hierarchies.</span></span> <span data-ttu-id="350d9-137">无论层次结构分配是通过 Excel UI 还是 JavaScript API Excel，此行为都是相同的。</span><span class="sxs-lookup"><span data-stu-id="350d9-137">This behavior is the same whether hierarchy assignment is done through the Excel UI or the Excel JavaScript APIs.</span></span>

<span data-ttu-id="350d9-138">行和列层次结构定义数据的分组方法。</span><span class="sxs-lookup"><span data-stu-id="350d9-138">Row and column hierarchies define how data will be grouped.</span></span> <span data-ttu-id="350d9-139">例如，服务器场的行 **层次结构将同** 一服务器场的所有数据集组合在一起。</span><span class="sxs-lookup"><span data-stu-id="350d9-139">For example, a row hierarchy of **Farms** will group together all the data sets from the same farm.</span></span> <span data-ttu-id="350d9-140">行和列层次结构之间的选择定义数据透视表的方向。</span><span class="sxs-lookup"><span data-stu-id="350d9-140">The choice between row and column hierarchy defines the orientation of the PivotTable.</span></span>

<span data-ttu-id="350d9-141">数据层次结构是基于行和列层次结构聚合的值。</span><span class="sxs-lookup"><span data-stu-id="350d9-141">Data hierarchies are the values to be aggregated based on the row and column hierarchies.</span></span> <span data-ttu-id="350d9-142">具有服务器场的行层次结构和数据层次结构"销售商品"的数据层次结构的数据透视表显示每个服务器场的所有不同客户 () 的默认总和。</span><span class="sxs-lookup"><span data-stu-id="350d9-142">A PivotTable with a row hierarchy of **Farms** and a data hierarchy of **Crates Sold Wholesale** shows the sum total (by default) of all the different fruits for each farm.</span></span>

<span data-ttu-id="350d9-143">筛选器层次结构基于该筛选类型中的值包含或排除数据透视表中的数据。</span><span class="sxs-lookup"><span data-stu-id="350d9-143">Filter hierarchies include or exclude data from the pivot based on values within that filtered type.</span></span> <span data-ttu-id="350d9-144">"分类"的筛选器 **层次结构（** 已选择 **"有机** "类型）只显示有机菜的数据。</span><span class="sxs-lookup"><span data-stu-id="350d9-144">A filter hierarchy of **Classification** with the type **Organic** selected only shows data for organic fruit.</span></span>

<span data-ttu-id="350d9-145">下面再次是数据透视表旁的服务器场数据。</span><span class="sxs-lookup"><span data-stu-id="350d9-145">Here is the farm data again, alongside a PivotTable.</span></span> <span data-ttu-id="350d9-146">数据透视表使用 **"服务器场**"和"类型"作为行层次结构，将"场中销售"和"出售的百货"作为数据层次结构 (其默认聚合函数为 sum) ，将 **Classification** 用作筛选器层次结构 (（选择 **"Organic**) "）。</span><span class="sxs-lookup"><span data-stu-id="350d9-146">The PivotTable is using **Farm** and **Type** as the row hierarchies, **Crates Sold at Farm** and **Crates Sold Wholesale** as the data hierarchies (with the default aggregation function of sum), and **Classification** as a filter hierarchy (with **Organic** selected).</span></span>

![数据透视表旁边包含一组包含行、数据和筛选器层次结构的新鲜菜销售数据。](../images/excel-pivot-table-and-data.png)

<span data-ttu-id="350d9-148">此数据透视表可以通过 JavaScript API 或 Excel UI 生成。</span><span class="sxs-lookup"><span data-stu-id="350d9-148">This PivotTable could be generated through the JavaScript API or through the Excel UI.</span></span> <span data-ttu-id="350d9-149">这两个选项都允许通过外接程序进一步操作。</span><span class="sxs-lookup"><span data-stu-id="350d9-149">Both options allow for further manipulation through add-ins.</span></span>

## <a name="create-a-pivottable"></a><span data-ttu-id="350d9-150">创建数据透视表</span><span class="sxs-lookup"><span data-stu-id="350d9-150">Create a PivotTable</span></span>

<span data-ttu-id="350d9-151">数据透视表需要名称、源和目标。</span><span class="sxs-lookup"><span data-stu-id="350d9-151">PivotTables need a name, source, and destination.</span></span> <span data-ttu-id="350d9-152">源可以是区域地址或表名称， (、 或 键入 `Range` `string` `Table`) 。</span><span class="sxs-lookup"><span data-stu-id="350d9-152">The source can be a range address or table name (passed as a `Range`, `string`, or `Table` type).</span></span> <span data-ttu-id="350d9-153">目标地址是一个范围 (给定为 `Range` 或 `string`) 。</span><span class="sxs-lookup"><span data-stu-id="350d9-153">The destination is a range address (given as either a `Range` or `string`).</span></span>
<span data-ttu-id="350d9-154">以下示例显示了各种数据透视表创建技术。</span><span class="sxs-lookup"><span data-stu-id="350d9-154">The following samples show various PivotTable creation techniques.</span></span>

### <a name="create-a-pivottable-with-range-addresses"></a><span data-ttu-id="350d9-155">创建具有区域地址的数据透视表</span><span class="sxs-lookup"><span data-stu-id="350d9-155">Create a PivotTable with range addresses</span></span>

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on the current worksheet at cell
    // A22 with data from the range A1:E21.
    context.workbook.worksheets.getActiveWorksheet().pivotTables.add(
      "Farm Sales", "A1:E21", "A22");

    return context.sync();
});
```

### <a name="create-a-pivottable-with-range-objects"></a><span data-ttu-id="350d9-156">创建包含 Range 对象的数据透视表</span><span class="sxs-lookup"><span data-stu-id="350d9-156">Create a PivotTable with Range objects</span></span>

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

### <a name="create-a-pivottable-at-the-workbook-level"></a><span data-ttu-id="350d9-157">在工作簿级别创建数据透视表</span><span class="sxs-lookup"><span data-stu-id="350d9-157">Create a PivotTable at the workbook level</span></span>

```js
Excel.run(function (context) {
    // Create a PivotTable named "Farm Sales" on a worksheet called "PivotWorksheet" at cell A2
    // the data is from the worksheet "DataWorksheet" across the range A1:E21.
    context.workbook.pivotTables.add(
        "Farm Sales", "DataWorksheet!A1:E21", "PivotWorksheet!A2");

    return context.sync();
});
```

## <a name="use-an-existing-pivottable"></a><span data-ttu-id="350d9-158">使用现有数据透视表</span><span class="sxs-lookup"><span data-stu-id="350d9-158">Use an existing PivotTable</span></span>

<span data-ttu-id="350d9-159">通过工作簿或单个工作表的数据透视表集合，也可以访问手动创建的数据透视表。</span><span class="sxs-lookup"><span data-stu-id="350d9-159">Manually created PivotTables are also accessible through the PivotTable collection of the workbook or of individual worksheets.</span></span> <span data-ttu-id="350d9-160">下面的代码从工作簿获取名为 **My Pivot** 的数据透视表。</span><span class="sxs-lookup"><span data-stu-id="350d9-160">The following code gets a PivotTable named **My Pivot** from the workbook.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.pivotTables.getItem("My Pivot");
    return context.sync();
});
```

## <a name="add-rows-and-columns-to-a-pivottable"></a><span data-ttu-id="350d9-161">向数据透视表添加行和列</span><span class="sxs-lookup"><span data-stu-id="350d9-161">Add rows and columns to a PivotTable</span></span>

<span data-ttu-id="350d9-162">行和列围绕这些字段的值透视数据。</span><span class="sxs-lookup"><span data-stu-id="350d9-162">Rows and columns pivot the data around those fields' values.</span></span>

<span data-ttu-id="350d9-163">添加 **"服务器场** "列可透视每个服务器场的所有销售额。</span><span class="sxs-lookup"><span data-stu-id="350d9-163">Adding the **Farm** column pivots all the sales around each farm.</span></span> <span data-ttu-id="350d9-164">添加 **"类型** "和" **分类** "行会进一步分解数据，这些数据基于所出售的菜以及它是否是有机的。</span><span class="sxs-lookup"><span data-stu-id="350d9-164">Adding the **Type** and **Classification** rows further breaks down the data based on what fruit was sold and whether it was organic or not.</span></span>

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

<span data-ttu-id="350d9-166">也可以只包含行或列的数据透视表。</span><span class="sxs-lookup"><span data-stu-id="350d9-166">You can also have a PivotTable with only rows or columns.</span></span>

```js
Excel.run(function (context) {
    var pivotTable = context.workbook.worksheets.getActiveWorksheet().pivotTables.getItem("Farm Sales");
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Farm"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Type"));
    pivotTable.rowHierarchies.add(pivotTable.hierarchies.getItem("Classification"));

    return context.sync();
});
```

## <a name="add-data-hierarchies-to-the-pivottable"></a><span data-ttu-id="350d9-167">向数据透视表添加数据层次结构</span><span class="sxs-lookup"><span data-stu-id="350d9-167">Add data hierarchies to the PivotTable</span></span>

<span data-ttu-id="350d9-168">数据层次结构使用基于行和列组合的信息填充数据透视表。</span><span class="sxs-lookup"><span data-stu-id="350d9-168">Data hierarchies fill the PivotTable with information to combine based on the rows and columns.</span></span> <span data-ttu-id="350d9-169">添加"场出售"和"出售 **的** 箱形"的数据层次结构会提供每一行和每一列的数据总和。</span><span class="sxs-lookup"><span data-stu-id="350d9-169">Adding the data hierarchies of **Crates Sold at Farm** and **Crates Sold Wholesale** gives sums of those figures for each row and column.</span></span>

<span data-ttu-id="350d9-170">在示例中 **，Farm** 和 **Type** 都是行，以销售为数据。</span><span class="sxs-lookup"><span data-stu-id="350d9-170">In the example, both **Farm** and **Type** are rows, with the crate sales as the data.</span></span>

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

## <a name="pivottable-layouts-and-getting-pivoted-data"></a><span data-ttu-id="350d9-172">数据透视表布局和获取透视数据</span><span class="sxs-lookup"><span data-stu-id="350d9-172">PivotTable layouts and getting pivoted data</span></span>

<span data-ttu-id="350d9-173">[PivotLayout](/javascript/api/excel/excel.pivotlayout)定义层次结构及其数据的位置。</span><span class="sxs-lookup"><span data-stu-id="350d9-173">A [PivotLayout](/javascript/api/excel/excel.pivotlayout) defines the placement of hierarchies and their data.</span></span> <span data-ttu-id="350d9-174">访问布局以确定存储数据的范围。</span><span class="sxs-lookup"><span data-stu-id="350d9-174">You access the layout to determine the ranges where data is stored.</span></span>

<span data-ttu-id="350d9-175">下图显示了哪些布局函数调用对应于数据透视表的哪些区域。</span><span class="sxs-lookup"><span data-stu-id="350d9-175">The following diagram shows which layout function calls correspond to which ranges of the PivotTable.</span></span>

![显示布局的 get range 函数返回数据透视表的哪些节的图表。](../images/excel-pivots-layout-breakdown.png)

### <a name="get-data-from-the-pivottable"></a><span data-ttu-id="350d9-177">从数据透视表获取数据</span><span class="sxs-lookup"><span data-stu-id="350d9-177">Get data from the PivotTable</span></span>

<span data-ttu-id="350d9-178">布局定义数据透视表在工作表中的显示方式。</span><span class="sxs-lookup"><span data-stu-id="350d9-178">The layout defines how the PivotTable is displayed in the worksheet.</span></span> <span data-ttu-id="350d9-179">这意味着对象 `PivotLayout` 控制用于数据透视表元素的范围。</span><span class="sxs-lookup"><span data-stu-id="350d9-179">This means the `PivotLayout` object controls the ranges used for PivotTable elements.</span></span> <span data-ttu-id="350d9-180">使用布局提供的范围获取数据透视表收集和聚合的数据。</span><span class="sxs-lookup"><span data-stu-id="350d9-180">Use the ranges provided by the layout to get data collected and aggregated by the PivotTable.</span></span> <span data-ttu-id="350d9-181">特别是，使用 `PivotLayout.getDataBodyRange` 访问数据透视表生成的数据。</span><span class="sxs-lookup"><span data-stu-id="350d9-181">In particular, use `PivotLayout.getDataBodyRange` to access the data produced by the PivotTable.</span></span>

<span data-ttu-id="350d9-182">下面的代码演示了如何获取数据透视表数据的最后一行，方法为在上一示例) 中浏览布局 (服务器场中"销售的箱形总和"和"销售的箱形总和"列。 </span><span class="sxs-lookup"><span data-stu-id="350d9-182">The following code demonstrates how to get the last row of the PivotTable data by going through the layout (the **Grand Total** of both the **Sum of Crates Sold at Farm** and **Sum of Crates Sold Wholesale** columns in the earlier example).</span></span> <span data-ttu-id="350d9-183">然后，这些值汇总在一起，最终总计显示在数据透视表数据透视表外部 (**E30** 单元格) 。</span><span class="sxs-lookup"><span data-stu-id="350d9-183">Those values are then summed together for a final total, which is displayed in cell **E30** (outside of the PivotTable).</span></span>

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

### <a name="layout-types"></a><span data-ttu-id="350d9-184">布局类型</span><span class="sxs-lookup"><span data-stu-id="350d9-184">Layout types</span></span>

<span data-ttu-id="350d9-185">数据透视表具有三种布局样式：精简、大纲和表格。</span><span class="sxs-lookup"><span data-stu-id="350d9-185">PivotTables have three layout styles: Compact, Outline, and Tabular.</span></span> <span data-ttu-id="350d9-186">我们已在之前的示例中看到简洁样式。</span><span class="sxs-lookup"><span data-stu-id="350d9-186">We've seen the compact style in the previous examples.</span></span>

<span data-ttu-id="350d9-187">下面的示例分别使用大纲样式和表格样式。</span><span class="sxs-lookup"><span data-stu-id="350d9-187">The following examples use the outline and tabular styles, respectively.</span></span> <span data-ttu-id="350d9-188">该代码示例演示如何在不同布局之间循环。</span><span class="sxs-lookup"><span data-stu-id="350d9-188">The code sample shows how to cycle between the different layouts.</span></span>

#### <a name="outline-layout"></a><span data-ttu-id="350d9-189">大纲布局</span><span class="sxs-lookup"><span data-stu-id="350d9-189">Outline layout</span></span>

![使用大纲布局的数据透视表。](../images/excel-pivots-outline-layout.png)

#### <a name="tabular-layout"></a><span data-ttu-id="350d9-191">表格布局</span><span class="sxs-lookup"><span data-stu-id="350d9-191">Tabular layout</span></span>

![使用表格布局的数据透视表。](../images/excel-pivots-tabular-layout.png)

#### <a name="pivotlayout-type-switch-code-sample"></a><span data-ttu-id="350d9-193">PivotLayout 类型开关代码示例</span><span class="sxs-lookup"><span data-stu-id="350d9-193">PivotLayout type switch code sample</span></span>

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

### <a name="other-pivotlayout-functions"></a><span data-ttu-id="350d9-194">其他 PivotLayout 函数</span><span class="sxs-lookup"><span data-stu-id="350d9-194">Other PivotLayout functions</span></span>

<span data-ttu-id="350d9-195">默认情况下，数据透视表会根据需要调整行和列大小。</span><span class="sxs-lookup"><span data-stu-id="350d9-195">By default, PivotTables adjust row and column sizes as needed.</span></span> <span data-ttu-id="350d9-196">此操作在数据透视表刷新时完成。</span><span class="sxs-lookup"><span data-stu-id="350d9-196">This is done when the PivotTable is refreshed.</span></span> <span data-ttu-id="350d9-197">`PivotLayout.autoFormat` 指定该行为。</span><span class="sxs-lookup"><span data-stu-id="350d9-197">`PivotLayout.autoFormat` specifies that behavior.</span></span> <span data-ttu-id="350d9-198">当 为 时，加载项进行的任何行或列大小更改将 `autoFormat` 持续存在 `false` 。</span><span class="sxs-lookup"><span data-stu-id="350d9-198">Any row or column size changes made by your add-in persist when `autoFormat` is `false`.</span></span> <span data-ttu-id="350d9-199">此外，数据透视表的默认设置在数据透视表中保留任何自定义 (如填充和字体) 。</span><span class="sxs-lookup"><span data-stu-id="350d9-199">Additionally, the default settings of a PivotTable keep any custom formatting in the PivotTable (such as fills and font changes).</span></span> <span data-ttu-id="350d9-200">设置为 `PivotLayout.preserveFormatting` `false` 以在刷新时应用默认格式。</span><span class="sxs-lookup"><span data-stu-id="350d9-200">Set `PivotLayout.preserveFormatting` to `false` to apply the default format when refreshed.</span></span>

<span data-ttu-id="350d9-201">还 `PivotLayout` 控制标题和总行设置、空数据单元格的显示方式以及 [替换文字](https://support.microsoft.com/topic/add-alternative-text-to-a-shape-picture-chart-smartart-graphic-or-other-object-44989b2a-903c-4d9a-b742-6a75b451c669) 选项。</span><span class="sxs-lookup"><span data-stu-id="350d9-201">A `PivotLayout` also controls header and total row settings, how empty data cells are displayed, and [alt text](https://support.microsoft.com/topic/add-alternative-text-to-a-shape-picture-chart-smartart-graphic-or-other-object-44989b2a-903c-4d9a-b742-6a75b451c669) options.</span></span> <span data-ttu-id="350d9-202">[PivotLayout](/javascript/api/excel/excel.pivotlayout)引用提供了这些功能的完整列表。</span><span class="sxs-lookup"><span data-stu-id="350d9-202">The [PivotLayout](/javascript/api/excel/excel.pivotlayout) reference provides a complete list of these features.</span></span>

<span data-ttu-id="350d9-203">下面的代码示例使空数据单元格显示字符串，将正文区域的格式设置为一致的水平对齐方式，并确保即使在数据透视表刷新后，格式更改 `"--"` 也保持不变。</span><span class="sxs-lookup"><span data-stu-id="350d9-203">The following code sample makes empty data cells display the string `"--"`, formats the body range to a consistent horizontal alignment, and ensures that the formatting changes remain even after the PivotTable is refreshed.</span></span>

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

## <a name="delete-a-pivottable"></a><span data-ttu-id="350d9-204">删除数据透视表</span><span class="sxs-lookup"><span data-stu-id="350d9-204">Delete a PivotTable</span></span>

<span data-ttu-id="350d9-205">使用数据透视表的名称删除数据透视表。</span><span class="sxs-lookup"><span data-stu-id="350d9-205">PivotTables are deleted by using their name.</span></span>

```js
Excel.run(function (context) {
    context.workbook.worksheets.getItem("Pivot").pivotTables.getItem("Farm Sales").delete();
    return context.sync();
});
```

## <a name="filter-a-pivottable"></a><span data-ttu-id="350d9-206">筛选数据透视表</span><span class="sxs-lookup"><span data-stu-id="350d9-206">Filter a PivotTable</span></span>

<span data-ttu-id="350d9-207">筛选数据透视表数据的主要方法是使用 PivotFilter。</span><span class="sxs-lookup"><span data-stu-id="350d9-207">The primary method for filtering PivotTable data is with PivotFilters.</span></span> <span data-ttu-id="350d9-208">切片器提供了一种不太灵活的备用筛选方法。</span><span class="sxs-lookup"><span data-stu-id="350d9-208">Slicers offer an alternate, less flexible filtering method.</span></span>

<span data-ttu-id="350d9-209">[PivotFilters](/javascript/api/excel/excel.pivotfilters)基于数据透视表的四个层次结构类别[](#hierarchies)筛选数据 (筛选器、列、行和) 。</span><span class="sxs-lookup"><span data-stu-id="350d9-209">[PivotFilters](/javascript/api/excel/excel.pivotfilters) filter data based on a PivotTable's four [hierarchy categories](#hierarchies) (filters, columns, rows, and values).</span></span> <span data-ttu-id="350d9-210">有四种类型的 PivotFilter，允许基于日历日期的筛选、字符串分析、数字比较和基于自定义输入的筛选。</span><span class="sxs-lookup"><span data-stu-id="350d9-210">There are four types of PivotFilters, allowing calendar date-based filtering, string parsing, number comparison, and filtering based on a custom input.</span></span>

<span data-ttu-id="350d9-211">[切片器](/javascript/api/excel/excel.slicer)可应用于数据透视表和常规Excel表。</span><span class="sxs-lookup"><span data-stu-id="350d9-211">[Slicers](/javascript/api/excel/excel.slicer) can be applied to both PivotTables and regular Excel tables.</span></span> <span data-ttu-id="350d9-212">应用于数据透视表时，切片器的功能与 [PivotManualFilter](#pivotmanualfilter) 类似，并允许基于自定义输入进行筛选。</span><span class="sxs-lookup"><span data-stu-id="350d9-212">When applied to a PivotTable, slicers function like a [PivotManualFilter](#pivotmanualfilter) and allow filtering based on a custom input.</span></span> <span data-ttu-id="350d9-213">与 PivotFilter 不同，切片器具有Excel [UI 组件](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)。</span><span class="sxs-lookup"><span data-stu-id="350d9-213">Unlike PivotFilters, slicers have an [Excel UI component](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d).</span></span> <span data-ttu-id="350d9-214">使用 `Slicer` 类，你可以创建此 UI 组件、管理筛选并控制其视觉外观。</span><span class="sxs-lookup"><span data-stu-id="350d9-214">With the `Slicer` class, you create this UI component, manage filtering, and control its visual appearance.</span></span>

### <a name="filter-with-pivotfilters"></a><span data-ttu-id="350d9-215">使用 PivotFilter 进行筛选</span><span class="sxs-lookup"><span data-stu-id="350d9-215">Filter with PivotFilters</span></span>

<span data-ttu-id="350d9-216">[PivotFilters](/javascript/api/excel/excel.pivotfilters)允许您基于四个层次结构类别筛选数据透视表[](#hierarchies)数据 (筛选器、列、行和) 。</span><span class="sxs-lookup"><span data-stu-id="350d9-216">[PivotFilters](/javascript/api/excel/excel.pivotfilters) allow you to filter PivotTable data based on the four [hierarchy categories](#hierarchies) (filters, columns, rows, and values).</span></span> <span data-ttu-id="350d9-217">在数据透视表对象模型中， `PivotFilters` 应用到透视 [字段](/javascript/api/excel/excel.pivotfield)，并且 `PivotField` 每个都可以分配一个或多个 `PivotFilters` 。</span><span class="sxs-lookup"><span data-stu-id="350d9-217">In the PivotTable object model, `PivotFilters` are applied to a [PivotField](/javascript/api/excel/excel.pivotfield), and each `PivotField` can have one or more assigned `PivotFilters`.</span></span> <span data-ttu-id="350d9-218">若要将 PivotFilter 应用于透视字段，必须将字段对应的 [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) 分配给层次结构类别。</span><span class="sxs-lookup"><span data-stu-id="350d9-218">To apply PivotFilters to a PivotField, the field's corresponding [PivotHierarchy](/javascript/api/excel/excel.pivothierarchy) must be assigned to a hierarchy category.</span></span>

#### <a name="types-of-pivotfilters"></a><span data-ttu-id="350d9-219">PivotFilter 的类型</span><span class="sxs-lookup"><span data-stu-id="350d9-219">Types of PivotFilters</span></span>

| <span data-ttu-id="350d9-220">筛选器类型</span><span class="sxs-lookup"><span data-stu-id="350d9-220">Filter type</span></span> | <span data-ttu-id="350d9-221">筛选目的</span><span class="sxs-lookup"><span data-stu-id="350d9-221">Filter purpose</span></span> | <span data-ttu-id="350d9-222">Excel JavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="350d9-222">Excel JavaScript API reference</span></span> |
|:--- |:--- |:--- |
| <span data-ttu-id="350d9-223">DateFilter</span><span class="sxs-lookup"><span data-stu-id="350d9-223">DateFilter</span></span> | <span data-ttu-id="350d9-224">基于日历日期的筛选。</span><span class="sxs-lookup"><span data-stu-id="350d9-224">Calendar date-based filtering.</span></span> | [<span data-ttu-id="350d9-225">PivotDateFilter</span><span class="sxs-lookup"><span data-stu-id="350d9-225">PivotDateFilter</span></span>](/javascript/api/excel/excel.pivotdatefilter) |
| <span data-ttu-id="350d9-226">LabelFilter</span><span class="sxs-lookup"><span data-stu-id="350d9-226">LabelFilter</span></span> | <span data-ttu-id="350d9-227">文本比较筛选。</span><span class="sxs-lookup"><span data-stu-id="350d9-227">Text comparison filtering.</span></span> | [<span data-ttu-id="350d9-228">PivotLabelFilter</span><span class="sxs-lookup"><span data-stu-id="350d9-228">PivotLabelFilter</span></span>](/javascript/api/excel/excel.pivotlabelfilter) |
| <span data-ttu-id="350d9-229">ManualFilter</span><span class="sxs-lookup"><span data-stu-id="350d9-229">ManualFilter</span></span> | <span data-ttu-id="350d9-230">自定义输入筛选。</span><span class="sxs-lookup"><span data-stu-id="350d9-230">Custom input filtering.</span></span> | [<span data-ttu-id="350d9-231">PivotManualFilter</span><span class="sxs-lookup"><span data-stu-id="350d9-231">PivotManualFilter</span></span>](/javascript/api/excel/excel.pivotmanualfilter) |
| <span data-ttu-id="350d9-232">ValueFilter</span><span class="sxs-lookup"><span data-stu-id="350d9-232">ValueFilter</span></span> | <span data-ttu-id="350d9-233">数字比较筛选。</span><span class="sxs-lookup"><span data-stu-id="350d9-233">Number comparison filtering.</span></span> | [<span data-ttu-id="350d9-234">PivotValueFilter</span><span class="sxs-lookup"><span data-stu-id="350d9-234">PivotValueFilter</span></span>](/javascript/api/excel/excel.pivotvaluefilter) |

#### <a name="create-a-pivotfilter"></a><span data-ttu-id="350d9-235">创建 PivotFilter</span><span class="sxs-lookup"><span data-stu-id="350d9-235">Create a PivotFilter</span></span>

<span data-ttu-id="350d9-236">若要使用数据透视表等 (`Pivot*Filter` 数据透视表 `PivotDateFilter`) ，请对透视字段应用 [筛选器](/javascript/api/excel/excel.pivotfield)。</span><span class="sxs-lookup"><span data-stu-id="350d9-236">To filter PivotTable data with a `Pivot*Filter` (such as a `PivotDateFilter`), apply the filter to a [PivotField](/javascript/api/excel/excel.pivotfield).</span></span> <span data-ttu-id="350d9-237">以下四个代码示例显示如何使用四种类型的 PivotFilter。</span><span class="sxs-lookup"><span data-stu-id="350d9-237">The following four code samples show how to use each of the four types of PivotFilters.</span></span>

##### <a name="pivotdatefilter"></a><span data-ttu-id="350d9-238">PivotDateFilter</span><span class="sxs-lookup"><span data-stu-id="350d9-238">PivotDateFilter</span></span>

<span data-ttu-id="350d9-239">第一个代码示例将 [PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) 应用于 **Date Updated** PivotField，隐藏 **2020-08-01 之前的任何数据**。</span><span class="sxs-lookup"><span data-stu-id="350d9-239">The first code sample applies a [PivotDateFilter](/javascript/api/excel/excel.pivotdatefilter) to the **Date Updated** PivotField, hiding any data prior to **2020-08-01**.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="350d9-240">`Pivot*Filter`无法将 应用于透视字段，除非将该字段的 PivotHierarchy 分配给层次结构类别。</span><span class="sxs-lookup"><span data-stu-id="350d9-240">A `Pivot*Filter` can't be applied to a PivotField unless that field's PivotHierarchy is assigned to a hierarchy category.</span></span> <span data-ttu-id="350d9-241">在下面的代码示例中，必须先将 添加到数据透视表的类别中，然后才能 `dateHierarchy` `rowHierarchies` 用于筛选。</span><span class="sxs-lookup"><span data-stu-id="350d9-241">In the following code sample, the `dateHierarchy` must be added to the PivotTable's `rowHierarchies` category before it can be used for filtering.</span></span>

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
> <span data-ttu-id="350d9-242">以下三个代码段仅显示特定于筛选器的摘要，而不是完整 `Excel.run` 调用。</span><span class="sxs-lookup"><span data-stu-id="350d9-242">The following three code snippets only display filter-specific excerpts, instead of full `Excel.run` calls.</span></span>

##### <a name="pivotlabelfilter"></a><span data-ttu-id="350d9-243">PivotLabelFilter</span><span class="sxs-lookup"><span data-stu-id="350d9-243">PivotLabelFilter</span></span>

<span data-ttu-id="350d9-244">第二个代码段演示如何将 [PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) 应用到 **Type** PivotField，使用 属性排除以 `LabelFilterCondition.beginsWith` 字母 L 开头 **的标签**。</span><span class="sxs-lookup"><span data-stu-id="350d9-244">The second code snippet demonstrates how to apply a [PivotLabelFilter](/javascript/api/excel/excel.pivotlabelfilter) to the **Type** PivotField, using the `LabelFilterCondition.beginsWith` property to exclude labels that start with the letter **L**.</span></span>

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

##### <a name="pivotmanualfilter"></a><span data-ttu-id="350d9-245">PivotManualFilter</span><span class="sxs-lookup"><span data-stu-id="350d9-245">PivotManualFilter</span></span>

<span data-ttu-id="350d9-246">第三个代码段使用 [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) 将手动筛选器应用于 **Classification** 字段，以筛选出不包含 **分类 Organic 的数据**。</span><span class="sxs-lookup"><span data-stu-id="350d9-246">The third code snippet applies a manual filter with [PivotManualFilter](/javascript/api/excel/excel.pivotmanualfilter) to the the **Classification** field, filtering out data that doesn't include the classification **Organic**.</span></span>

```js
    // Apply a manual filter to include only a specific PivotItem (the string "Organic").
    var filterField = classHierarchy.fields.getItem("Classification");
    var manualFilter = { selectedItems: ["Organic"] };
    filterField.applyFilter({ manualFilter: manualFilter });
```

##### <a name="pivotvaluefilter"></a><span data-ttu-id="350d9-247">PivotValueFilter</span><span class="sxs-lookup"><span data-stu-id="350d9-247">PivotValueFilter</span></span>

<span data-ttu-id="350d9-248">若要比较数字，请将值筛选器与 [PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter)一同使用，如最终代码片段中所示。</span><span class="sxs-lookup"><span data-stu-id="350d9-248">To compare numbers, use a value filter with [PivotValueFilter](/javascript/api/excel/excel.pivotvaluefilter), as shown in the final code snippet.</span></span> <span data-ttu-id="350d9-249">比较服务器场透视字段的数据与"已出售的 Crate PivotField"数据透视表中的数据，仅包括销售的箱总和超过 `PivotValueFilter` **值 500 的服务器场**。  </span><span class="sxs-lookup"><span data-stu-id="350d9-249">The `PivotValueFilter` compares the data in the **Farm** PivotField to the data in the **Crates Sold Wholesale** PivotField, including only farms whose sum of crates sold exceeds the value **500**.</span></span>

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

#### <a name="remove-pivotfilters"></a><span data-ttu-id="350d9-250">删除 PivotFilter</span><span class="sxs-lookup"><span data-stu-id="350d9-250">Remove PivotFilters</span></span>

<span data-ttu-id="350d9-251">若要删除所有 PivotFilter，请对各个 `clearAllFilters` 透视字段应用该方法，如下面的代码示例所示。</span><span class="sxs-lookup"><span data-stu-id="350d9-251">To remove all PivotFilters, apply the `clearAllFilters` method to each PivotField, as shown in the following code sample.</span></span>

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

### <a name="filter-with-slicers"></a><span data-ttu-id="350d9-252">使用切片器筛选</span><span class="sxs-lookup"><span data-stu-id="350d9-252">Filter with slicers</span></span>

<span data-ttu-id="350d9-253">[切片器](/javascript/api/excel/excel.slicer)允许从数据透视表或Excel筛选数据。</span><span class="sxs-lookup"><span data-stu-id="350d9-253">[Slicers](/javascript/api/excel/excel.slicer) allow data to be filtered from an Excel PivotTable or table.</span></span> <span data-ttu-id="350d9-254">切片器使用指定列或透视字段的值筛选相应的行。</span><span class="sxs-lookup"><span data-stu-id="350d9-254">A slicer uses values from a specified column or PivotField to filter corresponding rows.</span></span> <span data-ttu-id="350d9-255">这些值存储为 中的 [SlicerItem](/javascript/api/excel/excel.sliceritem) 对象 `Slicer` 。</span><span class="sxs-lookup"><span data-stu-id="350d9-255">These values are stored as [SlicerItem](/javascript/api/excel/excel.sliceritem) objects in the `Slicer`.</span></span> <span data-ttu-id="350d9-256">加载项可以调整这些筛选器，就像用户 (UI Excel[一](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)) 。</span><span class="sxs-lookup"><span data-stu-id="350d9-256">Your add-in can adjust these filters, as can users ([through the Excel UI](https://support.office.com/article/Use-slicers-to-filter-data-249f966b-a9d5-4b0f-b31a-12651785d29d)).</span></span> <span data-ttu-id="350d9-257">切片器位于绘图层中工作表的顶部，如以下屏幕截图所示。</span><span class="sxs-lookup"><span data-stu-id="350d9-257">The slicer sits on top of the worksheet in the drawing layer, as shown in the following screenshot.</span></span>

![筛选数据透视表上的数据的切片器。](../images/excel-slicer.png)

> [!NOTE]
> <span data-ttu-id="350d9-259">本节中介绍的技术侧重于如何使用连接到数据透视表的切片器。</span><span class="sxs-lookup"><span data-stu-id="350d9-259">The techniques described in this section focus on how to use slicers connected to PivotTables.</span></span> <span data-ttu-id="350d9-260">相同的技术也适用于使用连接到表的切片器。</span><span class="sxs-lookup"><span data-stu-id="350d9-260">The same techniques also apply to using slicers connected to tables.</span></span>

#### <a name="create-a-slicer"></a><span data-ttu-id="350d9-261">创建切片器</span><span class="sxs-lookup"><span data-stu-id="350d9-261">Create a slicer</span></span>

<span data-ttu-id="350d9-262">可以使用 方法在工作簿或工作表中 `Workbook.slicers.add` 创建切片 `Worksheet.slicers.add` 器。</span><span class="sxs-lookup"><span data-stu-id="350d9-262">You can create a slicer in a workbook or worksheet by using the `Workbook.slicers.add` method or `Worksheet.slicers.add` method.</span></span> <span data-ttu-id="350d9-263">这样做会将切片器添加到指定 或 对象的[SlicerCollection。](/javascript/api/excel/excel.slicercollection) `Workbook` `Worksheet`</span><span class="sxs-lookup"><span data-stu-id="350d9-263">Doing so adds a slicer to the [SlicerCollection](/javascript/api/excel/excel.slicercollection) of the specified `Workbook` or `Worksheet` object.</span></span> <span data-ttu-id="350d9-264">`SlicerCollection.add`方法具有三个参数：</span><span class="sxs-lookup"><span data-stu-id="350d9-264">The `SlicerCollection.add` method has three parameters:</span></span>

- <span data-ttu-id="350d9-265">`slicerSource`：新切片器所基于的数据源。</span><span class="sxs-lookup"><span data-stu-id="350d9-265">`slicerSource`: The data source on which the new slicer is based.</span></span> <span data-ttu-id="350d9-266">它可以是 `PivotTable` 、 `Table` 或 字符串，表示 或 的名称或 `PivotTable` `Table` ID。</span><span class="sxs-lookup"><span data-stu-id="350d9-266">It can be a `PivotTable`, `Table`, or string representing the name or ID of a `PivotTable` or `Table`.</span></span>
- <span data-ttu-id="350d9-267">`sourceField`：数据源中要筛选的字段。</span><span class="sxs-lookup"><span data-stu-id="350d9-267">`sourceField`: The field in the data source by which to filter.</span></span> <span data-ttu-id="350d9-268">它可以是 `PivotField` 、 `TableColumn` 或 字符串，表示 或 的名称或 `PivotField` `TableColumn` ID。</span><span class="sxs-lookup"><span data-stu-id="350d9-268">It can be a `PivotField`, `TableColumn`, or string representing the name or ID of a `PivotField` or `TableColumn`.</span></span>
- <span data-ttu-id="350d9-269">`slicerDestination`：将在其中新建切片器的工作表。</span><span class="sxs-lookup"><span data-stu-id="350d9-269">`slicerDestination`: The worksheet where the new slicer will be created.</span></span> <span data-ttu-id="350d9-270">它可以是 `Worksheet` 对象，或者是 的名称或 `Worksheet` ID。</span><span class="sxs-lookup"><span data-stu-id="350d9-270">It can be a `Worksheet` object or the name or ID of a `Worksheet`.</span></span> <span data-ttu-id="350d9-271">通过 访问 时， `SlicerCollection` 不需要此参数 `Worksheet.slicers` 。</span><span class="sxs-lookup"><span data-stu-id="350d9-271">This parameter is unnecessary when the `SlicerCollection` is accessed through `Worksheet.slicers`.</span></span> <span data-ttu-id="350d9-272">在这种情况下，集合的工作表用作目标。</span><span class="sxs-lookup"><span data-stu-id="350d9-272">In this case, the collection's worksheet is used as the destination.</span></span>

<span data-ttu-id="350d9-273">下面的代码示例向透视工作表 **添加新切片器** 。</span><span class="sxs-lookup"><span data-stu-id="350d9-273">The following code sample adds a new slicer to the **Pivot** worksheet.</span></span> <span data-ttu-id="350d9-274">切片器的来源是场 **销售** 数据透视表，并且使用 **Type 数据进行** 筛选。</span><span class="sxs-lookup"><span data-stu-id="350d9-274">The slicer's source is the **Farm Sales** PivotTable and filters using the **Type** data.</span></span> <span data-ttu-id="350d9-275">该切片器也命名为 **"切片器** "，供将来参考。</span><span class="sxs-lookup"><span data-stu-id="350d9-275">The slicer is also named **Fruit Slicer** for future reference.</span></span>

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

#### <a name="filter-items-with-a-slicer"></a><span data-ttu-id="350d9-276">使用切片器筛选项</span><span class="sxs-lookup"><span data-stu-id="350d9-276">Filter items with a slicer</span></span>

<span data-ttu-id="350d9-277">切片器使用 中的项筛选数据透视表 `sourceField` 。</span><span class="sxs-lookup"><span data-stu-id="350d9-277">The slicer filters the PivotTable with items from the `sourceField`.</span></span> <span data-ttu-id="350d9-278">`Slicer.selectItems`方法设置保留在切片器中的项。</span><span class="sxs-lookup"><span data-stu-id="350d9-278">The `Slicer.selectItems` method sets the items that remain in the slicer.</span></span> <span data-ttu-id="350d9-279">这些项目作为 传递给 方法， `string[]` 表示项的键。</span><span class="sxs-lookup"><span data-stu-id="350d9-279">These items are passed to the method as a `string[]`, representing the keys of the items.</span></span> <span data-ttu-id="350d9-280">包含这些项目的任何行都保留在数据透视表的聚合中。</span><span class="sxs-lookup"><span data-stu-id="350d9-280">Any rows containing those items remain in the PivotTable's aggregation.</span></span> <span data-ttu-id="350d9-281">后续调用 `selectItems` ，用于将列表设置为这些调用中指定的键。</span><span class="sxs-lookup"><span data-stu-id="350d9-281">Subsequent calls to `selectItems` set the list to the keys specified in those calls.</span></span>

> [!NOTE]
> <span data-ttu-id="350d9-282">`Slicer.selectItems`如果传递的项不在数据源中，则 `InvalidArgument` 会引发错误。</span><span class="sxs-lookup"><span data-stu-id="350d9-282">If `Slicer.selectItems` is passed an item that's not in the data source, an `InvalidArgument` error is thrown.</span></span> <span data-ttu-id="350d9-283">可通过 属性（即 `Slicer.slicerItems` [SlicerItemCollection ）验证内容](/javascript/api/excel/excel.sliceritemcollection)。</span><span class="sxs-lookup"><span data-stu-id="350d9-283">The contents can be verified through the `Slicer.slicerItems` property, which is a [SlicerItemCollection](/javascript/api/excel/excel.sliceritemcollection).</span></span>

<span data-ttu-id="350d9-284">下面的代码示例显示了为切片器选择的三个项目：**橙色**、**橙色\*\*\*\*和橙色**。</span><span class="sxs-lookup"><span data-stu-id="350d9-284">The following code sample shows three items being selected for the slicer: **Lemon**, **Lime**, and **Orange**.</span></span>

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    // Anything other than the following three values will be filtered out of the PivotTable for display and aggregation.
    slicer.selectItems(["Lemon", "Lime", "Orange"]);
    return context.sync();
});
```

<span data-ttu-id="350d9-285">若要从切片器中删除所有筛选器，请使用 `Slicer.clearFilters` 方法，如以下示例所示。</span><span class="sxs-lookup"><span data-stu-id="350d9-285">To remove all filters from the slicer, use the `Slicer.clearFilters` method, as shown in the following sample.</span></span>

```js
Excel.run(function (context) {
    var slicer = context.workbook.slicers.getItem("Fruit Slicer");
    slicer.clearFilters();
    return context.sync();
});
```

#### <a name="style-and-format-a-slicer"></a><span data-ttu-id="350d9-286">设置切片器样式和格式</span><span class="sxs-lookup"><span data-stu-id="350d9-286">Style and format a slicer</span></span>

<span data-ttu-id="350d9-287">加载项可以通过属性调整切片器显示 `Slicer` 设置。</span><span class="sxs-lookup"><span data-stu-id="350d9-287">You add-in can adjust a slicer's display settings through `Slicer` properties.</span></span> <span data-ttu-id="350d9-288">下面的代码示例将样式设置为 **SlicerStyleLight6**，将切片器顶部的文本设置为 **"菜** 类型"，将切片器放在绘图层上 **位置 (395，15) ，** 将切片器的大小设置为 **135x150** 像素。</span><span class="sxs-lookup"><span data-stu-id="350d9-288">The following code sample sets the style to **SlicerStyleLight6**, sets the text at the top of the slicer to **Fruit Types**, places the slicer at the position **(395, 15)** on the drawing layer, and sets the slicer's size to **135x150** pixels.</span></span>

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

#### <a name="delete-a-slicer"></a><span data-ttu-id="350d9-289">删除切片器</span><span class="sxs-lookup"><span data-stu-id="350d9-289">Delete a slicer</span></span>

<span data-ttu-id="350d9-290">若要删除切片器，请调用 `Slicer.delete` 方法。</span><span class="sxs-lookup"><span data-stu-id="350d9-290">To delete a slicer, call the `Slicer.delete` method.</span></span> <span data-ttu-id="350d9-291">下面的代码示例从当前工作表中删除第一个切片器。</span><span class="sxs-lookup"><span data-stu-id="350d9-291">The following code sample deletes the first slicer from the current worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.slicers.getItemAt(0).delete();
    return context.sync();
});
```

## <a name="change-aggregation-function"></a><span data-ttu-id="350d9-292">更改聚合函数</span><span class="sxs-lookup"><span data-stu-id="350d9-292">Change aggregation function</span></span>

<span data-ttu-id="350d9-293">数据层次结构聚合了它们的值。</span><span class="sxs-lookup"><span data-stu-id="350d9-293">Data hierarchies have their values aggregated.</span></span> <span data-ttu-id="350d9-294">对于数字数据集，默认情况下这是一个和。</span><span class="sxs-lookup"><span data-stu-id="350d9-294">For datasets of numbers, this is a sum by default.</span></span> <span data-ttu-id="350d9-295">属性 `summarizeBy` 根据 [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) 类型定义此行为。</span><span class="sxs-lookup"><span data-stu-id="350d9-295">The `summarizeBy` property defines this behavior based on an [AggregationFunction](/javascript/api/excel/excel.aggregationfunction) type.</span></span>

<span data-ttu-id="350d9-296">当前支持的聚合函数类型是 `Sum` 、和 `Count` (`Average` `Max` `Min` `Product` `CountNumbers` `StandardDeviation` `StandardDeviationP` `Variance` `VarianceP` `Automatic` 默认) 。</span><span class="sxs-lookup"><span data-stu-id="350d9-296">The currently supported aggregation function types are `Sum`, `Count`, `Average`, `Max`, `Min`, `Product`, `CountNumbers`, `StandardDeviation`, `StandardDeviationP`, `Variance`, `VarianceP`, and `Automatic` (the default).</span></span>

<span data-ttu-id="350d9-297">下面的代码示例将聚合更改为数据的平均值。</span><span class="sxs-lookup"><span data-stu-id="350d9-297">The following code samples changes the aggregation to be averages of the data.</span></span>

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

## <a name="change-calculations-with-a-showasrule"></a><span data-ttu-id="350d9-298">使用 ShowAsRule 更改计算</span><span class="sxs-lookup"><span data-stu-id="350d9-298">Change calculations with a ShowAsRule</span></span>

<span data-ttu-id="350d9-299">默认情况下，数据透视表独立聚合其行和列层次结构的数据。</span><span class="sxs-lookup"><span data-stu-id="350d9-299">PivotTables, by default, aggregate the data of their row and column hierarchies independently.</span></span> <span data-ttu-id="350d9-300">[ShowAsRule](/javascript/api/excel/excel.showasrule)根据数据透视表中的其他项将数据层次结构更改为输出值。</span><span class="sxs-lookup"><span data-stu-id="350d9-300">A [ShowAsRule](/javascript/api/excel/excel.showasrule) changes the data hierarchy to output values based on other items in the PivotTable.</span></span>

<span data-ttu-id="350d9-301">对象 `ShowAsRule` 有三个属性：</span><span class="sxs-lookup"><span data-stu-id="350d9-301">The `ShowAsRule` object has three properties:</span></span>

- <span data-ttu-id="350d9-302">`calculation`：应用于数据层次结构的相对计算类型 (默认值为 `none`) 。</span><span class="sxs-lookup"><span data-stu-id="350d9-302">`calculation`: The type of relative calculation to apply to the data hierarchy (the default is `none`).</span></span>
- <span data-ttu-id="350d9-303">`baseField`： [在应用](/javascript/api/excel/excel.pivotfield) 计算之前，层次结构中包含基本数据的透视字段。</span><span class="sxs-lookup"><span data-stu-id="350d9-303">`baseField`: The [PivotField](/javascript/api/excel/excel.pivotfield) in the hierarchy containing the base data before the calculation is applied.</span></span> <span data-ttu-id="350d9-304">由于Excel数据透视表具有到字段的一对一的层次结构映射，因此您将使用相同的名称访问层次结构和字段。</span><span class="sxs-lookup"><span data-stu-id="350d9-304">Since Excel PivotTables have a one-to-one mapping of hierarchy to field, you'll use the same name to access both the hierarchy and the field.</span></span>
- <span data-ttu-id="350d9-305">`baseItem`：单个 [PivotItem](/javascript/api/excel/excel.pivotitem) 与基于计算类型的基字段的值进行比较。</span><span class="sxs-lookup"><span data-stu-id="350d9-305">`baseItem`: The individual [PivotItem](/javascript/api/excel/excel.pivotitem) compared against the values of the base fields based on the calculation type.</span></span> <span data-ttu-id="350d9-306">并非所有计算都需要此字段。</span><span class="sxs-lookup"><span data-stu-id="350d9-306">Not all calculations require this field.</span></span>

<span data-ttu-id="350d9-307">下面的示例将场中销售的箱值总 **和数据层次结构** 的计算结果设定为列总计的百分比。</span><span class="sxs-lookup"><span data-stu-id="350d9-307">The following example sets the calculation on the **Sum of Crates Sold at Farm** data hierarchy to be a percentage of the column total.</span></span>
<span data-ttu-id="350d9-308">我们仍希望粒度扩展到树状类型级别，因此我们将使用 **Type** 行层次结构及其基础字段。</span><span class="sxs-lookup"><span data-stu-id="350d9-308">We still want the granularity to extend to the fruit type level, so we'll use the **Type** row hierarchy and its underlying field.</span></span>
<span data-ttu-id="350d9-309">此示例还将 **Farm** 作为第一行层次结构，因此服务器场的总条目也显示每个服务器场负责生成百分比。</span><span class="sxs-lookup"><span data-stu-id="350d9-309">The example also has **Farm** as the first row hierarchy, so the farm total entries display the percentage each farm is responsible for producing as well.</span></span>

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

<span data-ttu-id="350d9-311">前面的示例将相对于单个行层次结构的字段的计算设置为列。</span><span class="sxs-lookup"><span data-stu-id="350d9-311">The previous example set the calculation to the column, relative to the field of an individual row hierarchy.</span></span> <span data-ttu-id="350d9-312">当计算与单个项目相关时，请使用 `baseItem` 属性。</span><span class="sxs-lookup"><span data-stu-id="350d9-312">When the calculation relates to an individual item, use the `baseItem` property.</span></span>

<span data-ttu-id="350d9-313">以下示例显示了 `differenceFrom` 计算。</span><span class="sxs-lookup"><span data-stu-id="350d9-313">The following example shows the `differenceFrom` calculation.</span></span> <span data-ttu-id="350d9-314">它显示服务器场包含销售数据层次结构条目相对于服务器场 **的条目的区别**。</span><span class="sxs-lookup"><span data-stu-id="350d9-314">It displays the difference of the farm crate sales data hierarchy entries relative to those of **A Farms**.</span></span>
<span data-ttu-id="350d9-315">为 Farm，因此我们将看到其他服务器场之间的差异，以及每种类型的类似树 (Type 的细目也是此示例中的行层次结构 `baseField`) 。  </span><span class="sxs-lookup"><span data-stu-id="350d9-315">The `baseField` is **Farm**, so we see the differences between the other farms, as well as breakdowns for each type of like fruit (**Type** is also a row hierarchy in this example).</span></span>

![一个数据透视表，显示"A Farms"和其他服务器场之间的菜品销售差异。](../images/excel-pivots-showas-differencefrom.png)

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

## <a name="change-hierarchy-names"></a><span data-ttu-id="350d9-319">更改层次结构名称</span><span class="sxs-lookup"><span data-stu-id="350d9-319">Change hierarchy names</span></span>

<span data-ttu-id="350d9-320">层次结构字段是可编辑的。</span><span class="sxs-lookup"><span data-stu-id="350d9-320">Hierarchy fields are editable.</span></span> <span data-ttu-id="350d9-321">以下代码演示如何更改两个数据层次结构的显示名称。</span><span class="sxs-lookup"><span data-stu-id="350d9-321">The following code demonstrates how to change the displayed names of two data hierarchies.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="350d9-322">另请参阅</span><span class="sxs-lookup"><span data-stu-id="350d9-322">See also</span></span>

- [<span data-ttu-id="350d9-323">Excel 加载项中的 Word JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="350d9-323">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="350d9-324">ExcelJavaScript API 参考</span><span class="sxs-lookup"><span data-stu-id="350d9-324">Excel JavaScript API Reference</span></span>](/javascript/api/excel)
