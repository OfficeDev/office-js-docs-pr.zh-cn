---
title: 使用 Excel JavaScript API 处理公式引用单元格
description: 了解如何使用 Excel JavaScript API 检索公式引用单元格。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 0d21ae411615a22873a0f4dda185984f6191ac8e
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652796"
---
# <a name="get-formula-precedents-using-the-excel-javascript-api"></a><span data-ttu-id="a2cbd-103">使用 Excel JavaScript API 获取公式引用单元格</span><span class="sxs-lookup"><span data-stu-id="a2cbd-103">Get formula precedents using the Excel JavaScript API</span></span>

<span data-ttu-id="a2cbd-104">本文提供使用 Excel JavaScript API 检索公式引用单元格的代码示例。</span><span class="sxs-lookup"><span data-stu-id="a2cbd-104">This article provides a code sample that retrieves formula precedents using the Excel JavaScript API.</span></span> <span data-ttu-id="a2cbd-105">有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel.Range 类](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="a2cbd-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="get-formula-precedents"></a><span data-ttu-id="a2cbd-106">获取公式引用单元格</span><span class="sxs-lookup"><span data-stu-id="a2cbd-106">Get formula precedents</span></span>

<span data-ttu-id="a2cbd-107">Excel 公式通常引用其他单元格。</span><span class="sxs-lookup"><span data-stu-id="a2cbd-107">An Excel formula often refers to other cells.</span></span> <span data-ttu-id="a2cbd-108">当单元格向公式提供数据时，它称为公式"precedent"。</span><span class="sxs-lookup"><span data-stu-id="a2cbd-108">When a cell provides data to a formula, it is known as a formula "precedent".</span></span> <span data-ttu-id="a2cbd-109">若要了解有关与单元格之间的关系相关的 Excel 功能，请参阅 [显示公式和单元格之间的关系](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507)。</span><span class="sxs-lookup"><span data-stu-id="a2cbd-109">To learn more about Excel features related to relationships between cells, see [Display the relationships between formulas and cells](https://support.microsoft.com/office/display-the-relationships-between-formulas-and-cells-a59bef2b-3701-46bf-8ff1-d3518771d507).</span></span> 

<span data-ttu-id="a2cbd-110">使用 [Range.getDirectPrecedents，](/javascript/api/excel/excel.range#getdirectprecedents--)加载项可以定位公式的直接引用单元格。</span><span class="sxs-lookup"><span data-stu-id="a2cbd-110">With [Range.getDirectPrecedents](/javascript/api/excel/excel.range#getdirectprecedents--), your add-in can locate a formula's direct precedent cells.</span></span> <span data-ttu-id="a2cbd-111">`Range.getDirectPrecedents` 返回 `WorkbookRangeAreas` 一个对象。</span><span class="sxs-lookup"><span data-stu-id="a2cbd-111">`Range.getDirectPrecedents` returns a `WorkbookRangeAreas` object.</span></span> <span data-ttu-id="a2cbd-112">此对象包含工作簿中所有引用单元格的地址。</span><span class="sxs-lookup"><span data-stu-id="a2cbd-112">This object contains the addresses of all the precedents in the workbook.</span></span> <span data-ttu-id="a2cbd-113">对于每个包含 `RangeAreas` 至少一个公式引用单元格的工作表，它都有一个单独的对象。</span><span class="sxs-lookup"><span data-stu-id="a2cbd-113">It has a separate `RangeAreas` object for each worksheet containing at least one formula precedent.</span></span> <span data-ttu-id="a2cbd-114">有关使用对象的信息，请参阅在 Excel 加载项中同时处理 `RangeAreas` [多个区域](excel-add-ins-multiple-ranges.md)。</span><span class="sxs-lookup"><span data-stu-id="a2cbd-114">For more information on working with the `RangeAreas` object, see [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

<span data-ttu-id="a2cbd-115">在 Excel UI 中 **，"追踪引用** 单元格"按钮绘制从引用单元格到所选公式的箭头。</span><span class="sxs-lookup"><span data-stu-id="a2cbd-115">In the Excel UI, the **Trace Precedents** button draws an arrow from precedent cells to the selected formula.</span></span> <span data-ttu-id="a2cbd-116">与 Excel UI 按钮不同， `getDirectPrecedents` 该方法不绘制箭头。</span><span class="sxs-lookup"><span data-stu-id="a2cbd-116">Unlike the Excel UI button, the `getDirectPrecedents` method does not draw arrows.</span></span> 

> [!IMPORTANT]
> <span data-ttu-id="a2cbd-117">`getDirectPrecedents`方法无法跨工作簿检索引用单元格。</span><span class="sxs-lookup"><span data-stu-id="a2cbd-117">The `getDirectPrecedents` method can't retrieve precedent cells across workbooks.</span></span> 

<span data-ttu-id="a2cbd-118">下面的代码示例获取活动区域的直接引用单元格，然后将这些引用单元格的背景色更改为黄色。</span><span class="sxs-lookup"><span data-stu-id="a2cbd-118">The following code sample gets the direct precedents for the active range and then changes the background color of those precedent cells to yellow.</span></span> 

> [!NOTE]
> <span data-ttu-id="a2cbd-119">活动区域必须包含一个公式，该公式引用同一工作簿中的其他单元格，使突出显示正常工作。</span><span class="sxs-lookup"><span data-stu-id="a2cbd-119">The active range must contain a formula that references other cells in the same workbook for the highlighting to work properly.</span></span> 

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
        })
        .then(context.sync);
}).catch(errorHandlerFunction);
```

## <a name="see-also"></a><span data-ttu-id="a2cbd-120">另请参阅</span><span class="sxs-lookup"><span data-stu-id="a2cbd-120">See also</span></span>

- [<span data-ttu-id="a2cbd-121">Excel 加载项中的 Word JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="a2cbd-121">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="a2cbd-122">使用 Excel JavaScript API 处理单元格</span><span class="sxs-lookup"><span data-stu-id="a2cbd-122">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="a2cbd-123"> 同时在 Excel 加载项中处理多个区域 </span><span class="sxs-lookup"><span data-stu-id="a2cbd-123">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
