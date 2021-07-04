---
title: 使用 JavaScript API 设置并Excel区域
description: 了解如何使用 Excel JavaScript API 设置和获取选定区域Excel JavaScript API。
ms.date: 07/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 623ba5c1b9e76151d4a2c4b169e655236b37e8c8
ms.sourcegitcommit: aa73ec6367eaf74399fbf8d6b7776d77895e9982
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/03/2021
ms.locfileid: "53290780"
---
# <a name="set-and-get-the-selected-range-using-the-excel-javascript-api"></a><span data-ttu-id="ba8ec-103">使用 JavaScript API 设置并Excel区域</span><span class="sxs-lookup"><span data-stu-id="ba8ec-103">Set and get the selected range using the Excel JavaScript API</span></span>

<span data-ttu-id="ba8ec-104">本文提供了使用 JavaScript API 设置和获取选定区域Excel示例。</span><span class="sxs-lookup"><span data-stu-id="ba8ec-104">This article provides code samples that set and get the selected range with the Excel JavaScript API.</span></span> <span data-ttu-id="ba8ec-105">有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel。Range 类](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="ba8ec-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="set-the-selected-range"></a><span data-ttu-id="ba8ec-106">设置所选区域</span><span class="sxs-lookup"><span data-stu-id="ba8ec-106">Set the selected range</span></span>

<span data-ttu-id="ba8ec-107">下面的代码示例选择活动工作表中的区域 **B2:E6**。</span><span class="sxs-lookup"><span data-stu-id="ba8ec-107">The following code sample selects the range **B2:E6** in the active worksheet.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    var range = sheet.getRange("B2:E6");

    range.select();

    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="selected-range-b2e6"></a><span data-ttu-id="ba8ec-108">选定的区域 B2:E6</span><span class="sxs-lookup"><span data-stu-id="ba8ec-108">Selected range B2:E6</span></span>

![选定区域Excel。](../images/excel-ranges-set-selection.png)

## <a name="get-the-selected-range"></a><span data-ttu-id="ba8ec-110">获取所选区域</span><span class="sxs-lookup"><span data-stu-id="ba8ec-110">Get the selected range</span></span>

<span data-ttu-id="ba8ec-111">下面的代码示例获取所选区域、加载其 `address` 属性，然后向控制台写入一条消息。</span><span class="sxs-lookup"><span data-stu-id="ba8ec-111">The following code sample gets the selected range, loads its `address` property, and writes a message to the console.</span></span>

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

## <a name="select-the-edge-of-a-used-range"></a><span data-ttu-id="ba8ec-112">选择已用区域的边缘</span><span class="sxs-lookup"><span data-stu-id="ba8ec-112">Select the edge of a used range</span></span>

<span data-ttu-id="ba8ec-113">[Range.getRangeEdge](/javascript/api/excel/excel.range#getRangeEdge_direction__activeCell_)和[Range.getExtendedRange](/javascript/api/excel/excel.range#getExtendedRange_directionString__activeCell_)方法允许外接程序复制键盘选择快捷方式的行为，并基于当前所选区域选择已用区域的边缘。</span><span class="sxs-lookup"><span data-stu-id="ba8ec-113">The [Range.getRangeEdge](/javascript/api/excel/excel.range#getRangeEdge_direction__activeCell_) and [Range.getExtendedRange](/javascript/api/excel/excel.range#getExtendedRange_directionString__activeCell_) methods let your add-in replicate the behavior of the keyboard selection shortcuts, selecting the edge of the used range based on the currently selected range.</span></span> <span data-ttu-id="ba8ec-114">若要了解有关已用区域有关详细信息，请参阅 [获取已用区域](excel-add-ins-ranges-get.md#get-used-range)。</span><span class="sxs-lookup"><span data-stu-id="ba8ec-114">To learn more about used ranges, see [Get used range](excel-add-ins-ranges-get.md#get-used-range).</span></span>

<span data-ttu-id="ba8ec-115">在下面的屏幕截图中，使用的范围是每个单元格中具有值的表 **C5：F12**。</span><span class="sxs-lookup"><span data-stu-id="ba8ec-115">In the following screenshot, the used range is the table with values in each cell, **C5:F12**.</span></span> <span data-ttu-id="ba8ec-116">此表外部的空单元格位于已用区域之外。</span><span class="sxs-lookup"><span data-stu-id="ba8ec-116">The empty cells outside this table are outside the used range.</span></span>

![包含来自 C5：F12 的数据的Excel。](../images/excel-ranges-used-range.png)

### <a name="select-the-cell-at-the-edge-of-the-current-used-range"></a><span data-ttu-id="ba8ec-118">选择当前使用区域边缘的单元格</span><span class="sxs-lookup"><span data-stu-id="ba8ec-118">Select the cell at the edge of the current used range</span></span>

<span data-ttu-id="ba8ec-119">下面的代码示例演示如何使用 方法按向上方向选择当前使用区域最远边缘 `Range.getRangeEdge` 的单元格。</span><span class="sxs-lookup"><span data-stu-id="ba8ec-119">The following code sample shows how use the `Range.getRangeEdge` method to select the cell at the furthest edge of the current used range, in the direction up.</span></span> <span data-ttu-id="ba8ec-120">此操作与选择范围时使用 Ctrl+向上箭头键键盘快捷方式的结果匹配。</span><span class="sxs-lookup"><span data-stu-id="ba8ec-120">This action matches the result of using the Ctrl+Up arrow key keyboard shortcut while a range is selected.</span></span>

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

#### <a name="before-selecting-the-cell-at-the-edge-of-the-used-range"></a><span data-ttu-id="ba8ec-121">选择已用区域边缘的单元格之前</span><span class="sxs-lookup"><span data-stu-id="ba8ec-121">Before selecting the cell at the edge of the used range</span></span>

<span data-ttu-id="ba8ec-122">以下屏幕截图显示了已用区域以及已用区域内的选定区域。</span><span class="sxs-lookup"><span data-stu-id="ba8ec-122">The following screenshot shows a used range and a selected range within the used range.</span></span> <span data-ttu-id="ba8ec-123">使用的范围是一个表，其数据位于 **C5：F12。**</span><span class="sxs-lookup"><span data-stu-id="ba8ec-123">The used range is a table with data at **C5:F12**.</span></span> <span data-ttu-id="ba8ec-124">在此表中，选择区域 **D8：E9。**</span><span class="sxs-lookup"><span data-stu-id="ba8ec-124">Inside this table, the range **D8:E9** is selected.</span></span> <span data-ttu-id="ba8ec-125">此选择 *是运行 方法* 之前的状态 `Range.getRangeEdge` 。</span><span class="sxs-lookup"><span data-stu-id="ba8ec-125">This selection is the *before* state, prior to running the `Range.getRangeEdge` method.</span></span>

![包含来自 C5：F12 的数据的Excel。](../images/excel-ranges-used-range-d8-e9.png)

#### <a name="after-selecting-the-cell-at-the-edge-of-the-used-range"></a><span data-ttu-id="ba8ec-128">选择已用区域边缘的单元格后</span><span class="sxs-lookup"><span data-stu-id="ba8ec-128">After selecting the cell at the edge of the used range</span></span>

<span data-ttu-id="ba8ec-129">以下屏幕截图显示与上一屏幕截图相同的表，包含 **C5：F12 范围的数据**。</span><span class="sxs-lookup"><span data-stu-id="ba8ec-129">The following screenshot shows the same table as the preceding screenshot, with data in the range **C5:F12**.</span></span> <span data-ttu-id="ba8ec-130">在此表中，选择了区域 **D5。**</span><span class="sxs-lookup"><span data-stu-id="ba8ec-130">Inside this table, the range **D5** is selected.</span></span> <span data-ttu-id="ba8ec-131">此选择 *位于状态* 之后，运行方法以在向上方向选择已用区域边缘 `Range.getRangeEdge` 的单元格。</span><span class="sxs-lookup"><span data-stu-id="ba8ec-131">This selection is *after* state, after running the `Range.getRangeEdge` method to select the cell at the edge of the used range in the up direction.</span></span>

![包含来自 C5：F12 的数据的Excel。](../images/excel-ranges-used-range-d5.png)

### <a name="select-all-cells-from-current-range-to-furthest-edge-of-used-range"></a><span data-ttu-id="ba8ec-134">选择从当前区域到已用区域最远边缘的所有单元格</span><span class="sxs-lookup"><span data-stu-id="ba8ec-134">Select all cells from current range to furthest edge of used range</span></span>

<span data-ttu-id="ba8ec-135">下面的代码示例演示如何使用 方法按向下方向选择从当前所选区域到已用区域最远边缘 `Range.getExtendedRange` 的所有单元格。</span><span class="sxs-lookup"><span data-stu-id="ba8ec-135">The following code sample shows how use the `Range.getExtendedRange` method to to select all the cells from the currently selected range to the furthest edge of the used range, in the direction down.</span></span> <span data-ttu-id="ba8ec-136">此操作与选中区域时使用 Ctrl+Shift+向下箭头键键盘快捷方式的结果匹配。</span><span class="sxs-lookup"><span data-stu-id="ba8ec-136">This action matches the result of using the Ctrl+Shift+Down arrow key keyboard shortcut while a range is selected.</span></span>

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

#### <a name="before-selecting-all-the-cells-from-the-current-range-to-the-edge-of-the-used-range"></a><span data-ttu-id="ba8ec-137">选择当前区域到已用区域边缘的所有单元格之前</span><span class="sxs-lookup"><span data-stu-id="ba8ec-137">Before selecting all the cells from the current range to the edge of the used range</span></span>

<span data-ttu-id="ba8ec-138">以下屏幕截图显示了已用区域以及已用区域内的选定区域。</span><span class="sxs-lookup"><span data-stu-id="ba8ec-138">The following screenshot shows a used range and a selected range within the used range.</span></span> <span data-ttu-id="ba8ec-139">使用的范围是一个表，其数据位于 **C5：F12。**</span><span class="sxs-lookup"><span data-stu-id="ba8ec-139">The used range is a table with data at **C5:F12**.</span></span> <span data-ttu-id="ba8ec-140">在此表中，选择区域 **D8：E9。**</span><span class="sxs-lookup"><span data-stu-id="ba8ec-140">Inside this table, the range **D8:E9** is selected.</span></span> <span data-ttu-id="ba8ec-141">此选择 *是运行 方法* 之前的状态 `Range.getExtendedRange` 。</span><span class="sxs-lookup"><span data-stu-id="ba8ec-141">This selection is the *before* state, prior to running the `Range.getExtendedRange` method.</span></span>

![包含来自 C5：F12 的数据的Excel。](../images/excel-ranges-used-range-d8-e9.png)

#### <a name="after-selecting-all-the-cells-from-the-current-range-to-the-edge-of-the-used-range"></a><span data-ttu-id="ba8ec-144">选择从当前区域到已用区域边缘的所有单元格后</span><span class="sxs-lookup"><span data-stu-id="ba8ec-144">After selecting all the cells from the current range to the edge of the used range</span></span>

<span data-ttu-id="ba8ec-145">以下屏幕截图显示与上一屏幕截图相同的表，包含 **C5：F12 范围的数据**。</span><span class="sxs-lookup"><span data-stu-id="ba8ec-145">The following screenshot shows the same table as the preceding screenshot, with data in the range **C5:F12**.</span></span> <span data-ttu-id="ba8ec-146">在此表中，选择了区域 **D8：E12。**</span><span class="sxs-lookup"><span data-stu-id="ba8ec-146">Inside this table, the range **D8:E12** is selected.</span></span> <span data-ttu-id="ba8ec-147">在运行 *该方法以* 从当前区域到已用区域的边缘沿向下方向选择所有单元格之后，此选择处于 `Range.getExtendedRange` 之后状态。</span><span class="sxs-lookup"><span data-stu-id="ba8ec-147">This selection is *after* state, after running the `Range.getExtendedRange` method to select all the cells from the current range to the edge of the used range in the down direction.</span></span>

![包含来自 C5：F12 的数据的Excel。](../images/excel-ranges-used-range-d8-e12.png)

## <a name="see-also"></a><span data-ttu-id="ba8ec-150">另请参阅</span><span class="sxs-lookup"><span data-stu-id="ba8ec-150">See also</span></span>

- [<span data-ttu-id="ba8ec-151">Excel 加载项中的 Word JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="ba8ec-151">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="ba8ec-152">使用 JavaScript API Excel单元格</span><span class="sxs-lookup"><span data-stu-id="ba8ec-152">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="ba8ec-153">使用 JavaScript API 设置和获取区域Excel文本或公式</span><span class="sxs-lookup"><span data-stu-id="ba8ec-153">Set and get range values, text, or formulas using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-get-values.md)
- [<span data-ttu-id="ba8ec-154">使用 JavaScript API Excel区域格式</span><span class="sxs-lookup"><span data-stu-id="ba8ec-154">Set range format using the Excel JavaScript API</span></span>](excel-add-ins-ranges-set-format.md)
