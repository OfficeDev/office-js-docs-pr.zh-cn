---
title: 使用 JavaScript API 设置并Excel区域
description: 了解如何使用 Excel JavaScript API 设置和获取选定区域Excel JavaScript API。
ms.date: 06/22/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 9e4c31f165b39d45fac342cb85577ef737105472
ms.sourcegitcommit: ebb4a22a0bdeb5623c72b9494ebbce3909d0c90c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/25/2021
ms.locfileid: "53126720"
---
# <a name="set-and-get-the-selected-range-using-the-excel-javascript-api"></a><span data-ttu-id="8c454-103">使用 JavaScript API 设置并Excel区域</span><span class="sxs-lookup"><span data-stu-id="8c454-103">Set and get the selected range using the Excel JavaScript API</span></span>

<span data-ttu-id="8c454-104">本文提供了使用 JavaScript API 设置和获取选定区域Excel示例。</span><span class="sxs-lookup"><span data-stu-id="8c454-104">This article provides code samples that set and get the selected range with the Excel JavaScript API.</span></span> <span data-ttu-id="8c454-105">有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel。Range 类](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="8c454-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-the-selected-range"></a><span data-ttu-id="8c454-106">设置所选区域</span><span class="sxs-lookup"><span data-stu-id="8c454-106">Set the selected range</span></span>

<span data-ttu-id="8c454-107">下面的代码示例选择活动工作表中的区域 **B2:E6**。</span><span class="sxs-lookup"><span data-stu-id="8c454-107">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="selected-range-b2e6"></a><span data-ttu-id="8c454-108">选定的区域 B2:E6</span><span class="sxs-lookup"><span data-stu-id="8c454-108">Selected range B2:E6</span></span>

![选定区域Excel。](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="8c454-110">获取所选区域</span><span class="sxs-lookup"><span data-stu-id="8c454-110">Get the selected range</span></span>

<span data-ttu-id="8c454-111">下面的代码示例获取所选区域、加载其 `address` 属性，然后向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="8c454-111">The following code sample gets the selected range, loads its `address` property, and writes a message to the console.</span></span>

```js
Excel.run(function (context) {
    var range = context.workbook.getSelectedRange();
    range.load("address");

    return context.sync()
        .then(function () {
            console.log(`The address of the selected range is "${range.address}"`);
        });
}).catch(errorHandlerFunction);
```

## <a name="select-the-edge-of-a-used-range-online-only"></a><span data-ttu-id="8c454-112">选择已使用区域的边缘 (仅联机) </span><span class="sxs-lookup"><span data-stu-id="8c454-112">Select the edge of a used range (online-only)</span></span>

> [!NOTE]
> <span data-ttu-id="8c454-113">和 `Range.getRangeEdge` `Range.getExtendedRange` 方法当前仅在 ExcelApiOnline 1.1 中可用。</span><span class="sxs-lookup"><span data-stu-id="8c454-113">The `Range.getRangeEdge` and `Range.getExtendedRange` methods are currently only available in ExcelApiOnline 1.1.</span></span> <span data-ttu-id="8c454-114">若要了解更多信息，请参阅[Excel JavaScript API 仅联机要求集](../reference/requirement-sets/excel-api-online-requirement-set.md)。</span><span class="sxs-lookup"><span data-stu-id="8c454-114">To learn more, see [Excel JavaScript API online-only requirement set](../reference/requirement-sets/excel-api-online-requirement-set.md).</span></span>

<span data-ttu-id="8c454-115">[Range.getRangeEdge](/javascript/api/excel/excel.range#getRangeEdge_direction__activeCell_)和[Range.getExtendedRange](/javascript/api/excel/excel.range#getExtendedRange_directionString__activeCell_)方法允许外接程序复制键盘选择快捷方式的行为，并基于当前所选区域选择已用区域的边缘。</span><span class="sxs-lookup"><span data-stu-id="8c454-115">The [Range.getRangeEdge](/javascript/api/excel/excel.range#getRangeEdge_direction__activeCell_) and [Range.getExtendedRange](/javascript/api/excel/excel.range#getExtendedRange_directionString__activeCell_) methods let your add-in replicate the behavior of the keyboard selection shortcuts, selecting the edge of the used range based on the currently selected range.</span></span> <span data-ttu-id="8c454-116">若要了解有关已用区域有关详细信息，请参阅 [获取已用区域](excel-add-ins-ranges-get.md#get-used-range)。</span><span class="sxs-lookup"><span data-stu-id="8c454-116">To learn more about used ranges, see [Get used range](excel-add-ins-ranges-get.md#get-used-range).</span></span>

<span data-ttu-id="8c454-117">在下面的屏幕截图中，使用的范围是每个单元格中具有值的表 **C5：F12**。</span><span class="sxs-lookup"><span data-stu-id="8c454-117">In the following screenshot, the used range is the table with values in each cell, **C5:F12**.</span></span> <span data-ttu-id="8c454-118">此表外部的空单元格位于已用区域之外。</span><span class="sxs-lookup"><span data-stu-id="8c454-118">The empty cells outside this table are outside the used range.</span></span>

![包含来自 C5：F12 的数据的Excel。](../images/excel-ranges-used-range.png)

### <a name="select-the-cell-at-the-edge-of-the-current-used-range"></a><span data-ttu-id="8c454-120">选择当前使用区域边缘的单元格</span><span class="sxs-lookup"><span data-stu-id="8c454-120">Select the cell at the edge of the current used range</span></span>

<span data-ttu-id="8c454-121">下面的代码示例演示如何使用 方法按向上方向选择当前使用区域最远边缘 `Range.getRangeEdge` 的单元格。</span><span class="sxs-lookup"><span data-stu-id="8c454-121">The following code sample shows how use the `Range.getRangeEdge` method to select the cell at the furthest edge of the current used range, in the direction up.</span></span> <span data-ttu-id="8c454-122">此操作与选择范围时使用 Ctrl+向上箭头键键盘快捷方式的结果匹配。</span><span class="sxs-lookup"><span data-stu-id="8c454-122">This action matches the result of using the Ctrl+Up arrow key keyboard shortcut while a range is selected.</span></span>

```js
Excel.run(function (context) {
    // Get the selected range.
    var range = context.workbook.getSelectedRange();

    // Specify the direction with the `KeyboardDirection` enum.
    var direction = Excel.KeyboardDirection.up;

    // Get the active cell in the workbook.
    var activeCell = context.workbook.getActiveCell();

    // Get the top-most cell of the current used range.
    // This method acts like the Ctrl+Up arrow key keyboard shortcut while a range is selected.
    var rangeEdge = range.getRangeEdge(
      direction,
      activeCell
    );
    rangeEdge.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="before-selecting-the-cell-at-the-edge-of-the-used-range"></a><span data-ttu-id="8c454-123">选择已用区域边缘的单元格之前</span><span class="sxs-lookup"><span data-stu-id="8c454-123">Before selecting the cell at the edge of the used range</span></span>

<span data-ttu-id="8c454-124">以下屏幕截图显示了已用区域以及已用区域内的选定区域。</span><span class="sxs-lookup"><span data-stu-id="8c454-124">The following screenshot shows a used range and a selected range within the used range.</span></span> <span data-ttu-id="8c454-125">使用的范围是一个表，其数据位于 **C5：F12。**</span><span class="sxs-lookup"><span data-stu-id="8c454-125">The used range is a table with data at **C5:F12**.</span></span> <span data-ttu-id="8c454-126">在此表中，选择区域 **D8：E9。**</span><span class="sxs-lookup"><span data-stu-id="8c454-126">Inside this table, the range **D8:E9** is selected.</span></span> <span data-ttu-id="8c454-127">此选择 *是运行 方法* 之前的状态 `Range.getRangeEdge` 。</span><span class="sxs-lookup"><span data-stu-id="8c454-127">This selection is the *before* state, prior to running the `Range.getRangeEdge` method.</span></span>

![包含来自 C5：F12 的数据的Excel。](../images/excel-ranges-used-range-d8-e9.png)

#### <a name="after-selecting-the-cell-at-the-edge-of-the-used-range"></a><span data-ttu-id="8c454-130">选择已用区域边缘的单元格后</span><span class="sxs-lookup"><span data-stu-id="8c454-130">After selecting the cell at the edge of the used range</span></span>

<span data-ttu-id="8c454-131">以下屏幕截图显示与上一屏幕截图相同的表，包含 **C5：F12 范围的数据**。</span><span class="sxs-lookup"><span data-stu-id="8c454-131">The following screenshot shows the same table as the preceding screenshot, with data in the range **C5:F12**.</span></span> <span data-ttu-id="8c454-132">在此表中，选择了区域 **D5。**</span><span class="sxs-lookup"><span data-stu-id="8c454-132">Inside this table, the range **D5** is selected.</span></span> <span data-ttu-id="8c454-133">此选择 *位于状态* 之后，运行方法以在向上方向选择已用区域边缘 `Range.getRangeEdge` 的单元格。</span><span class="sxs-lookup"><span data-stu-id="8c454-133">This selection is *after* state, after running the `Range.getRangeEdge` method to select the cell at the edge of the used range in the up direction.</span></span>

![包含来自 C5：F12 的数据的Excel。](../images/excel-ranges-used-range-d5.png)

### <a name="select-all-cells-from-current-range-to-furthest-edge-of-used-range"></a><span data-ttu-id="8c454-136">选择从当前区域到已用区域最远边缘的所有单元格</span><span class="sxs-lookup"><span data-stu-id="8c454-136">Select all cells from current range to furthest edge of used range</span></span>

<span data-ttu-id="8c454-137">下面的代码示例演示如何使用 方法按向下方向选择从当前所选区域到已用区域最远边缘 `Range.getExtendedRange` 的所有单元格。</span><span class="sxs-lookup"><span data-stu-id="8c454-137">The following code sample shows how use the `Range.getExtendedRange` method to to select all the cells from the currently selected range to the furthest edge of the used range, in the direction down.</span></span> <span data-ttu-id="8c454-138">此操作与选中区域时使用 Ctrl+Shift+向下箭头键键盘快捷方式的结果匹配。</span><span class="sxs-lookup"><span data-stu-id="8c454-138">This action matches the result of using the Ctrl+Shift+Down arrow key keyboard shortcut while a range is selected.</span></span>

```js
Excel.run(function (context) {
    // Get the selected range.
    var range = context.workbook.getSelectedRange();

    // Specify the direction with the `KeyboardDirection` enum.
    var direction = Excel.KeyboardDirection.down;

    // Get the active cell in the workbook.
    var activeCell = context.workbook.getActiveCell();

    // Get all the cells from the currently selected range to the bottom-most edge of the used range.
    // This method acts like the Ctrl+Shift+Down arrow key keyboard shortcut while a range is selected.
    var extendedRange = range.getExtendedRange(
      direction,
      activeCell
    );
    extendedRange.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

#### <a name="before-selecting-all-the-cells-from-the-current-range-to-the-edge-of-the-used-range"></a><span data-ttu-id="8c454-139">选择当前区域到已用区域边缘的所有单元格之前</span><span class="sxs-lookup"><span data-stu-id="8c454-139">Before selecting all the cells from the current range to the edge of the used range</span></span>

<span data-ttu-id="8c454-140">以下屏幕截图显示了已用区域以及已用区域内的选定区域。</span><span class="sxs-lookup"><span data-stu-id="8c454-140">The following screenshot shows a used range and a selected range within the used range.</span></span> <span data-ttu-id="8c454-141">使用的范围是一个表，其数据位于 **C5：F12。**</span><span class="sxs-lookup"><span data-stu-id="8c454-141">The used range is a table with data at **C5:F12**.</span></span> <span data-ttu-id="8c454-142">在此表中，选择区域 **D8：E9。**</span><span class="sxs-lookup"><span data-stu-id="8c454-142">Inside this table, the range **D8:E9** is selected.</span></span> <span data-ttu-id="8c454-143">此选择 *是运行 方法* 之前的状态 `Range.getExtendedRange` 。</span><span class="sxs-lookup"><span data-stu-id="8c454-143">This selection is the *before* state, prior to running the `Range.getExtendedRange` method.</span></span>

![包含来自 C5：F12 的数据的Excel。](../images/excel-ranges-used-range-d8-e9.png)

#### <a name="after-selecting-all-the-cells-from-the-current-range-to-the-edge-of-the-used-range"></a><span data-ttu-id="8c454-146">选择从当前区域到已用区域边缘的所有单元格后</span><span class="sxs-lookup"><span data-stu-id="8c454-146">After selecting all the cells from the current range to the edge of the used range</span></span>

<span data-ttu-id="8c454-147">以下屏幕截图显示与上一屏幕截图相同的表，包含 **C5：F12 范围的数据**。</span><span class="sxs-lookup"><span data-stu-id="8c454-147">The following screenshot shows the same table as the preceding screenshot, with data in the range **C5:F12**.</span></span> <span data-ttu-id="8c454-148">在此表中，选择了区域 **D8：E12。**</span><span class="sxs-lookup"><span data-stu-id="8c454-148">Inside this table, the range **D8:E12** is selected.</span></span> <span data-ttu-id="8c454-149">在运行 *该方法以* 从当前区域到已用区域的边缘沿向下方向选择所有单元格之后，此选择处于 `Range.getExtendedRange` 之后状态。</span><span class="sxs-lookup"><span data-stu-id="8c454-149">This selection is *after* state, after running the `Range.getExtendedRange` method to select all the cells from the current range to the edge of the used range in the down direction.</span></span>

![包含来自 C5：F12 的数据的Excel。](../images/excel-ranges-used-range-d8-e12.png)

## <a name="see-also"></a><span data-ttu-id="8c454-152">另请参阅</span><span class="sxs-lookup"><span data-stu-id="8c454-152">See also</span></span>

- [<span data-ttu-id="8c454-153">Excel 加载项中的 Word JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="8c454-153">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="8c454-154">使用 JavaScript API Excel单元格</span><span class="sxs-lookup"><span data-stu-id="8c454-154">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="8c454-155">使用 JavaScript API 设置和获取区域Excel文本或公式</span><span class="sxs-lookup"><span data-stu-id="8c454-155">Set and get range values, text, or formulas using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get-values.md)
- [<span data-ttu-id="8c454-156">使用 JavaScript API Excel区域格式</span><span class="sxs-lookup"><span data-stu-id="8c454-156">Set range format using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-format.md)
