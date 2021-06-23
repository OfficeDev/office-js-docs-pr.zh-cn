---
title: 使用 JavaScript API Excel组范围
description: 了解如何将一个范围的行或列组合在一起，以使用 JavaScript API Excel大纲。
ms.date: 04/05/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 960a394a1467ec1fe55ff8dbf7b0a3f39fd355a5
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075717"
---
# <a name="group-ranges-for-an-outline-using-the-excel-javascript-api"></a><span data-ttu-id="4f9a8-103">使用 JavaScript API 的大纲Excel区域</span><span class="sxs-lookup"><span data-stu-id="4f9a8-103">Group ranges for an outline using the Excel JavaScript API</span></span>

<span data-ttu-id="4f9a8-104">本文提供了一个代码示例，演示如何使用 JavaScript API 对大纲Excel分组。</span><span class="sxs-lookup"><span data-stu-id="4f9a8-104">This article provides a code sample that shows how to group ranges for an outline using the Excel JavaScript API.</span></span> <span data-ttu-id="4f9a8-105">有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel。Range 类](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="4f9a8-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="group-rows-or-columns-of-a-range-for-an-outline"></a><span data-ttu-id="4f9a8-106">对分级显示区域行或列进行分组</span><span class="sxs-lookup"><span data-stu-id="4f9a8-106">Group rows or columns of a range for an outline</span></span>

<span data-ttu-id="4f9a8-107">可以将范围的行或列组合在一起以 [创建大纲](https://support.office.com/article/Outline-group-data-in-a-worksheet-08CE98C4-0063-4D42-8AC7-8278C49E9AFF)。</span><span class="sxs-lookup"><span data-stu-id="4f9a8-107">Rows or columns of a range can be grouped together to create an [outline](https://support.office.com/article/Outline-group-data-in-a-worksheet-08CE98C4-0063-4D42-8AC7-8278C49E9AFF).</span></span> <span data-ttu-id="4f9a8-108">可以折叠和展开这些组，以隐藏和显示相应的单元格。</span><span class="sxs-lookup"><span data-stu-id="4f9a8-108">These groups can be collapsed and expanded to hide and show the corresponding cells.</span></span> <span data-ttu-id="4f9a8-109">这使得快速分析首行数据变得更容易。</span><span class="sxs-lookup"><span data-stu-id="4f9a8-109">This makes quick analysis of top-line data easier.</span></span> <span data-ttu-id="4f9a8-110">使用 [Range.group](/javascript/api/excel/excel.range#group-groupoption-) 可创建这些大纲组。</span><span class="sxs-lookup"><span data-stu-id="4f9a8-110">Use [Range.group](/javascript/api/excel/excel.range#group-groupoption-) to make these outline groups.</span></span>

<span data-ttu-id="4f9a8-111">大纲可以具有层次结构，其中较小的组嵌套在较大的组下。</span><span class="sxs-lookup"><span data-stu-id="4f9a8-111">An outline can have a hierarchy, where smaller groups are nested under larger groups.</span></span> <span data-ttu-id="4f9a8-112">这允许在不同级别查看大纲。</span><span class="sxs-lookup"><span data-stu-id="4f9a8-112">This allows the outline to be viewed at different levels.</span></span> <span data-ttu-id="4f9a8-113">可以通过 [Worksheet.showOutlineLevels](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-) 方法以编程方式更改可见大纲级别。</span><span class="sxs-lookup"><span data-stu-id="4f9a8-113">Changing the visible outline level can be done programmatically through the [Worksheet.showOutlineLevels](/javascript/api/excel/excel.worksheet#showoutlinelevels-rowlevels--columnlevels-) method.</span></span> <span data-ttu-id="4f9a8-114">请注意，Excel仅支持八个级别的大纲组。</span><span class="sxs-lookup"><span data-stu-id="4f9a8-114">Note that Excel only supports eight levels of outline groups.</span></span>

<span data-ttu-id="4f9a8-115">下面的代码示例为行和列创建包含两个级别的组的大纲。</span><span class="sxs-lookup"><span data-stu-id="4f9a8-115">The following code sample creates an outline with two levels of groups for both the rows and columns.</span></span> <span data-ttu-id="4f9a8-116">后续图像显示该轮廓的分组。</span><span class="sxs-lookup"><span data-stu-id="4f9a8-116">The subsequent image shows the groupings of that outline.</span></span> <span data-ttu-id="4f9a8-117">在代码示例中，分组的范围不包括大纲控件的行或列 (此示例的"总计") 。</span><span class="sxs-lookup"><span data-stu-id="4f9a8-117">In the code sample, the ranges being grouped do not include the row or column of the outline control (the "Totals" for this example).</span></span> <span data-ttu-id="4f9a8-118">组定义要折叠的项，而不是控件的行或列。</span><span class="sxs-lookup"><span data-stu-id="4f9a8-118">A group defines what will be collapsed, not the row or column with the control.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Group the larger, main level. Note that the outline controls
    // will be on row 10, meaning 4-9 will collapse and expand.
    sheet.getRange("4:9").group(Excel.GroupOption.byRows);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on rows 6 and 9, meaning 4-5 and 7-8 will collapse and expand.
    sheet.getRange("4:5").group(Excel.GroupOption.byRows);
    sheet.getRange("7:8").group(Excel.GroupOption.byRows);

    // Group the larger, main level. Note that the outline controls
    // will be on column R, meaning C-Q will collapse and expand.
    sheet.getRange("C:Q").group(Excel.GroupOption.byColumns);

    // Group the smaller, sublevels. Note that the outline controls
    // will be on columns G, L, and R, meaning C-F, H-K, and M-P will collapse and expand.
    sheet.getRange("C:F").group(Excel.GroupOption.byColumns);
    sheet.getRange("H:K").group(Excel.GroupOption.byColumns);
    sheet.getRange("M:P").group(Excel.GroupOption.byColumns);
    return context.sync();
}).catch(errorHandlerFunction);
```

![具有两级、二维轮廓的范围。](../images/excel-outline.png)

## <a name="remove-grouping-from-rows-or-columns-of-a-range"></a><span data-ttu-id="4f9a8-120">从区域行或列中删除分组</span><span class="sxs-lookup"><span data-stu-id="4f9a8-120">Remove grouping from rows or columns of a range</span></span>

<span data-ttu-id="4f9a8-121">若要取消行或列组的分组，请使用 [Range.ungroup](/javascript/api/excel/excel.range#ungroup-groupoption-) 方法。</span><span class="sxs-lookup"><span data-stu-id="4f9a8-121">To ungroup a row or column group, use the [Range.ungroup](/javascript/api/excel/excel.range#ungroup-groupoption-) method.</span></span> <span data-ttu-id="4f9a8-122">这将从大纲中删除最外面的级别。</span><span class="sxs-lookup"><span data-stu-id="4f9a8-122">This removes the outermost level from the outline.</span></span> <span data-ttu-id="4f9a8-123">如果同一行或列类型的多个组位于指定范围内的同一级别，则所有这些组将取消分组。</span><span class="sxs-lookup"><span data-stu-id="4f9a8-123">If multiple groups of the same row or column type are at the same level within the specified range, all of those groups are ungrouped.</span></span>

## <a name="see-also"></a><span data-ttu-id="4f9a8-124">另请参阅</span><span class="sxs-lookup"><span data-stu-id="4f9a8-124">See also</span></span>

- [<span data-ttu-id="4f9a8-125">Excel 加载项中的 Word JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="4f9a8-125">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="4f9a8-126">使用 JavaScript API Excel单元格</span><span class="sxs-lookup"><span data-stu-id="4f9a8-126">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="4f9a8-127"> 同时在 Excel 加载项中处理多个区域 </span><span class="sxs-lookup"><span data-stu-id="4f9a8-127">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
