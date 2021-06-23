---
title: 使用 JavaScript API 处理公式引用Excel依赖项
description: 了解如何使用 JavaScript API Excel引用单元格和依赖项。
ms.date: 06/03/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 6021e383f02ca0de15210638b991dfe8b109ab63
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075794"
---
# <a name="get-formula-precedents-and-dependents-using-the-excel-javascript-api"></a><span data-ttu-id="9be8d-103">使用 JavaScript API 获取公式引用Excel依赖项</span><span class="sxs-lookup"><span data-stu-id="9be8d-103">Get formula precedents and dependents using the Excel JavaScript API</span></span>

<span data-ttu-id="9be8d-104">Excel公式通常引用其他单元格。</span><span class="sxs-lookup"><span data-stu-id="9be8d-104">Excel formulas often refer to other cells.</span></span> <span data-ttu-id="9be8d-105">这些跨单元格引用称为"引用单元格"和"从属单元格"。</span><span class="sxs-lookup"><span data-stu-id="9be8d-105">These cross-cell references are known as "precedents" and "dependents".</span></span> <span data-ttu-id="9be8d-106">引用单元格是向公式提供数据的单元格。</span><span class="sxs-lookup"><span data-stu-id="9be8d-106">A precedent is a cell that provides data to a formula.</span></span> <span data-ttu-id="9be8d-107">从属单元格是包含引用其他单元格的公式的单元格。</span><span class="sxs-lookup"><span data-stu-id="9be8d-107">A dependent is a cell that contains a formula that refers to other cells.</span></span> <span data-ttu-id="9be8d-108">若要了解有关与Excel关系相关的功能，请参阅显示[公式和单元格之间的关系](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507)。</span><span class="sxs-lookup"><span data-stu-id="9be8d-108">To learn more about Excel features related to relationships between cells, see [Display the relationships between formulas and cells](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507).</span></span>

<span data-ttu-id="9be8d-109">单元格可以有引用单元格，并且该引用单元格可能有自己的引用单元格。</span><span class="sxs-lookup"><span data-stu-id="9be8d-109">A cell may have a precedent cell, and that precedent cell may have its own precedent cells.</span></span> <span data-ttu-id="9be8d-110">"直接引用单元格"是此序列中前面的第一组单元格，类似于父子关系中父级的概念。</span><span class="sxs-lookup"><span data-stu-id="9be8d-110">A "direct precedent" is the first preceding group of cells in this sequence, similar to the concept of parents in a parent-child relationship.</span></span> <span data-ttu-id="9be8d-111">"直接从属"是序列中第一个从属单元格组，类似于父子关系中的子级。</span><span class="sxs-lookup"><span data-stu-id="9be8d-111">A "direct dependent" is the first dependent group of cells in a sequence, similar to children in a parent-child relationship.</span></span> <span data-ttu-id="9be8d-112">引用工作簿中其他单元格，但其关系不是父子关系的单元格不是直接从属单元格或直接引用单元格。</span><span class="sxs-lookup"><span data-stu-id="9be8d-112">Cells that refer to other cells in a workbook, but whose relationship is not a parent-child relationship, are not direct dependents or direct precedents.</span></span>

<span data-ttu-id="9be8d-113">本文提供的代码示例使用 JavaScript API 检索公式的直接引用Excel从属单元格。</span><span class="sxs-lookup"><span data-stu-id="9be8d-113">This article provides code samples that retrieve direct precedents and direct dependents of formulas using the Excel JavaScript API.</span></span> <span data-ttu-id="9be8d-114">有关对象支持的属性和方法的完整列表，请参阅 `Range` Range Object [ (JavaScript API for Excel) 。 ](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="9be8d-114">For the complete list of properties and methods that the `Range` object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).</span></span>

## <a name="get-the-direct-precedents-of-a-formula"></a><span data-ttu-id="9be8d-115">获取公式的直接引用单元格</span><span class="sxs-lookup"><span data-stu-id="9be8d-115">Get the direct precedents of a formula</span></span>

<span data-ttu-id="9be8d-116">使用 [Range.getDirectPrecedents](/javascript/api/excel/excel.range#getdirectprecedents--)查找公式的直接引用单元格。</span><span class="sxs-lookup"><span data-stu-id="9be8d-116">Locate a formula's direct precedent cells with [Range.getDirectPrecedents](/javascript/api/excel/excel.range#getdirectprecedents--).</span></span> <span data-ttu-id="9be8d-117">`Range.getDirectPrecedents` 返回 `WorkbookRangeAreas` 一个对象。</span><span class="sxs-lookup"><span data-stu-id="9be8d-117">`Range.getDirectPrecedents` returns a `WorkbookRangeAreas` object.</span></span> <span data-ttu-id="9be8d-118">此对象包含工作簿中所有直接引用单元格的地址。</span><span class="sxs-lookup"><span data-stu-id="9be8d-118">This object contains the addresses of all the direct precedents in the workbook.</span></span> <span data-ttu-id="9be8d-119">对于每个包含 `RangeAreas` 至少一个公式引用单元格的工作表，它都有一个单独的对象。</span><span class="sxs-lookup"><span data-stu-id="9be8d-119">It has a separate `RangeAreas` object for each worksheet containing at least one formula precedent.</span></span> <span data-ttu-id="9be8d-120">有关使用对象的信息，请参阅在加载项中同时Excel `RangeAreas` [多个区域](excel-add-ins-multiple-ranges.md)。</span><span class="sxs-lookup"><span data-stu-id="9be8d-120">For more information on working with the `RangeAreas` object, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

<span data-ttu-id="9be8d-121">以下屏幕截图显示了在"跟踪引用单元格"UI 中选择"追踪引用Excel的结果。 </span><span class="sxs-lookup"><span data-stu-id="9be8d-121">The following screenshot shows the result of selecting the **Trace Precedents** button in the Excel UI.</span></span> <span data-ttu-id="9be8d-122">此按钮绘制从引用单元格到选定单元格的箭头。</span><span class="sxs-lookup"><span data-stu-id="9be8d-122">This button draws an arrow from precedent cells to the selected cell.</span></span> <span data-ttu-id="9be8d-123">选定的单元格 **E3** 包含公式"=C3 \* D3"，因此 **C3** 和 **D3 都是** 引用单元格。</span><span class="sxs-lookup"><span data-stu-id="9be8d-123">The selected cell, **E3**, contains the formula "=C3 \* D3", so both **C3** and **D3** are precedent cells.</span></span> <span data-ttu-id="9be8d-124">与 Excel UI 按钮不同， `getDirectPrecedents` 该方法不绘制箭头。</span><span class="sxs-lookup"><span data-stu-id="9be8d-124">Unlike the Excel UI button, the `getDirectPrecedents` method does not draw arrows.</span></span>

![箭头跟踪活动 UI 中的引用单元格Excel单元格。](../images/excel-ranges-trace-precedents.png)

> [!IMPORTANT]
> <span data-ttu-id="9be8d-126">`getDirectPrecedents`方法无法跨工作簿检索引用单元格。</span><span class="sxs-lookup"><span data-stu-id="9be8d-126">The `getDirectPrecedents` method can't retrieve precedent cells across workbooks.</span></span>

<span data-ttu-id="9be8d-127">下面的代码示例获取活动区域的直接引用单元格，然后将这些引用单元格的背景色更改为黄色。</span><span class="sxs-lookup"><span data-stu-id="9be8d-127">The following code sample gets the direct precedents for the active range and then changes the background color of those precedent cells to yellow.</span></span>

```js
Excel.run(function (context) {
    // Precedents are cells that provide data to the selected formula.
    var range = context.workbook.getActiveCell();
    var directPrecedents = range.getDirectPrecedents();
    range.load("address");
    directPrecedents.areas.load("address");
    
    return context.sync()
        .then(function () {
            console.log(`Direct precedent cells of ${range.address}:`);

            // Use the direct precedents API to loop through precedents of the active cell.
            for (var i = 0; i < directPrecedents.areas.items.length; i++) {
              // Highlight and print out the address of each precedent cell.
              directPrecedents.areas.items[i].format.fill.color = "Yellow";
              console.log(`  ${directPrecedents.areas.items[i].address}`);
            }
        });
}).catch(errorHandlerFunction);
```

## <a name="get-the-direct-dependents-of-a-formula-preview"></a><span data-ttu-id="9be8d-128">获取公式的直接依赖项 (预览) </span><span class="sxs-lookup"><span data-stu-id="9be8d-128">Get the direct dependents of a formula (preview)</span></span>

> [!NOTE]
> <span data-ttu-id="9be8d-129">`Range.getDirectDependents`方法当前仅适用于公共预览版。</span><span class="sxs-lookup"><span data-stu-id="9be8d-129">The `Range.getDirectDependents` method is currently only available in public preview.</span></span> [!INCLUDE [Information about using preview APIs](../includes/using-excel-preview-apis.md)]
> 

<span data-ttu-id="9be8d-130">使用 [Range.getDirectDependents 查找公式的直接从属单元格](/javascript/api/excel/excel.range#getDirectDependents__)。</span><span class="sxs-lookup"><span data-stu-id="9be8d-130">Locate a formula's direct dependent cells with [Range.getDirectDependents](/javascript/api/excel/excel.range#getDirectDependents__).</span></span> <span data-ttu-id="9be8d-131">与 `Range.getDirectPrecedents` 类似 `Range.getDirectDependents` ，也返回 `WorkbookRangeAreas` 对象。</span><span class="sxs-lookup"><span data-stu-id="9be8d-131">Like `Range.getDirectPrecedents`, `Range.getDirectDependents` also returns a `WorkbookRangeAreas` object.</span></span> <span data-ttu-id="9be8d-132">此对象包含工作簿中所有直接依赖项的地址。</span><span class="sxs-lookup"><span data-stu-id="9be8d-132">This object contains the addresses of all the direct dependents in the workbook.</span></span> <span data-ttu-id="9be8d-133">对于每个包含至少一个依赖公式的工作表，它都有 `RangeAreas` 一个单独的对象。</span><span class="sxs-lookup"><span data-stu-id="9be8d-133">It has a separate `RangeAreas` object for each worksheet containing at least one formula dependent.</span></span> <span data-ttu-id="9be8d-134">有关使用对象的信息，请参阅在加载项中同时Excel `RangeAreas` [多个区域](excel-add-ins-multiple-ranges.md)。</span><span class="sxs-lookup"><span data-stu-id="9be8d-134">For more information on working with the `RangeAreas` object, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

<span data-ttu-id="9be8d-135">以下屏幕截图显示了在自定义 UI 中选择"跟踪从属 **项**"Excel的结果。</span><span class="sxs-lookup"><span data-stu-id="9be8d-135">The following screenshot shows the result of selecting the **Trace Dependents** button in the Excel UI.</span></span> <span data-ttu-id="9be8d-136">此按钮绘制从从属单元格到选定单元格的箭头。</span><span class="sxs-lookup"><span data-stu-id="9be8d-136">This button draws an arrow from dependent cells to the selected cell.</span></span> <span data-ttu-id="9be8d-137">选定的单元格 **D3** 将单元格 **E3** 作为从属单元格。</span><span class="sxs-lookup"><span data-stu-id="9be8d-137">The selected cell, **D3**, has cell **E3** as a dependent.</span></span> <span data-ttu-id="9be8d-138">**E3** 包含公式"=C3 \* D3"。</span><span class="sxs-lookup"><span data-stu-id="9be8d-138">**E3** contains the formula "=C3 \* D3".</span></span> <span data-ttu-id="9be8d-139">与 Excel UI 按钮不同， `getDirectDependents` 该方法不绘制箭头。</span><span class="sxs-lookup"><span data-stu-id="9be8d-139">Unlike the Excel UI button, the `getDirectDependents` method does not draw arrows.</span></span>

![箭头跟踪 UI 中的Excel单元格。](../images/excel-ranges-trace-dependents.png)

> [!IMPORTANT]
> <span data-ttu-id="9be8d-141">`getDirectDependents`方法无法跨工作簿检索从属单元格。</span><span class="sxs-lookup"><span data-stu-id="9be8d-141">The `getDirectDependents` method can't retrieve dependent cells across workbooks.</span></span>

<span data-ttu-id="9be8d-142">下面的代码示例获取活动区域的直接从属单元格，然后将这些从属单元格的背景色更改为黄色。</span><span class="sxs-lookup"><span data-stu-id="9be8d-142">The following code sample gets the direct dependents for the active range and then changes the background color of those dependent cells to yellow.</span></span>

```js
Excel.run(function (context) {
    // Direct dependents are cells that contain formulas that refer to other cells.
    var range = context.workbook.getActiveCell();
    var directDependents = range.getDirectDependents();
    range.load("address");
    directDependents.areas.load("address");
    
    return context.sync()
        .then(function () {
            console.log(`Direct dependent cells of ${range.address}:`);
    
            // Use the direct dependents API to loop through direct dependents of the active cell.
            for (var i = 0; i < directDependents.areas.items.length; i++) {
              // Highlight and print the address of each dependent cell.
              directDependents.areas.items[i].format.fill.color = "Yellow";
              console.log(`  ${directDependents.areas.items[i].address}`);
            }
        });
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a><span data-ttu-id="9be8d-143">另请参阅</span><span class="sxs-lookup"><span data-stu-id="9be8d-143">See also</span></span>

- [<span data-ttu-id="9be8d-144">Excel 加载项中的 Word JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="9be8d-144">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="9be8d-145">使用 JavaScript API Excel单元格</span><span class="sxs-lookup"><span data-stu-id="9be8d-145">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="9be8d-146"> 同时在 Excel 加载项中处理多个区域 </span><span class="sxs-lookup"><span data-stu-id="9be8d-146">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
