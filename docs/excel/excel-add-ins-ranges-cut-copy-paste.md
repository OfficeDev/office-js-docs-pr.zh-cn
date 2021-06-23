---
title: 使用 JavaScript API Excel、复制和粘贴区域
description: 了解如何使用 JavaScript API 剪切、复制和粘贴Excel区域。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 2112702110b72e0020ed72090ce495abb3ff5366
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075822"
---
# <a name="cut-copy-and-paste-ranges-using-the-excel-javascript-api"></a><span data-ttu-id="6e57c-103">使用 JavaScript API Excel、复制和粘贴区域</span><span class="sxs-lookup"><span data-stu-id="6e57c-103">Cut, copy, and paste ranges using the Excel JavaScript API</span></span>

<span data-ttu-id="6e57c-104">本文提供使用 JavaScript API 剪切、复制和粘贴区域Excel示例。</span><span class="sxs-lookup"><span data-stu-id="6e57c-104">This article provides code samples that cut, copy, and paste ranges using the Excel JavaScript API.</span></span> <span data-ttu-id="6e57c-105">有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel。Range 类](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="6e57c-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

[!include[Excel cells and ranges note](../includes/note-excel-cells-and-ranges.md)]

## <a name="copy-and-paste"></a><span data-ttu-id="6e57c-106">Copy and paste</span><span class="sxs-lookup"><span data-stu-id="6e57c-106">Copy and paste</span></span>

<span data-ttu-id="6e57c-107">[Range.copyFrom](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-)方法复制该 **UI** **的** 复制Excel粘贴操作。</span><span class="sxs-lookup"><span data-stu-id="6e57c-107">The [Range.copyFrom](/javascript/api/excel/excel.range#copyfrom-sourcerange--copytype--skipblanks--transpose-) method replicates the **Copy** and **Paste** actions of the Excel UI.</span></span> <span data-ttu-id="6e57c-108">目标为 `Range` 所 `copyFrom` 调用的对象。</span><span class="sxs-lookup"><span data-stu-id="6e57c-108">The destination is the `Range` object that `copyFrom` is called on.</span></span> <span data-ttu-id="6e57c-109">将要复制的源作为一个范围或一个表示范围的字符串地址进行传递。</span><span class="sxs-lookup"><span data-stu-id="6e57c-109">The source to be copied is passed as a range or a string address representing a range.</span></span>

<span data-ttu-id="6e57c-110">以下代码示例将数据从“A1:E1”复制到“G1”开始的范围（粘贴到“G1:K1”结束）。</span><span class="sxs-lookup"><span data-stu-id="6e57c-110">The following code sample copies the data from **A1:E1** into the range starting at **G1** (which ends up pasting into **G1:K1**).</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy everything from "A1:E1" into "G1" and the cells afterwards ("G1:K1")
    sheet.getRange("G1").copyFrom("A1:E1");
    return context.sync();
}).catch(errorHandlerFunction);
```

<span data-ttu-id="6e57c-111">`Range.copyFrom` 具有三个可选参数。</span><span class="sxs-lookup"><span data-stu-id="6e57c-111">`Range.copyFrom` has three optional parameters.</span></span>

```TypeScript
copyFrom(sourceRange: Range | RangeAreas | string, copyType?: Excel.RangeCopyType, skipBlanks?: boolean, transpose?: boolean): void;
```

<span data-ttu-id="6e57c-112">`copyType` 指定将哪些数据从源复制到目标。</span><span class="sxs-lookup"><span data-stu-id="6e57c-112">`copyType` specifies what data gets copied from the source to the destination.</span></span>

- <span data-ttu-id="6e57c-113">`Excel.RangeCopyType.formulas` 传输源单元格中的公式，并保留这些公式区域的相对位置。</span><span class="sxs-lookup"><span data-stu-id="6e57c-113">`Excel.RangeCopyType.formulas` transfers the formulas in the source cells and preserves the relative positioning of those formulas' ranges.</span></span> <span data-ttu-id="6e57c-114">将原样复制任何非公式条目。</span><span class="sxs-lookup"><span data-stu-id="6e57c-114">Any non-formula entries are copied as-is.</span></span>
- <span data-ttu-id="6e57c-115">`Excel.RangeCopyType.values` 复制数据值，如果是公式，则复制公式的结果。</span><span class="sxs-lookup"><span data-stu-id="6e57c-115">`Excel.RangeCopyType.values` copies the data values and, in the case of formulas, the result of the formula.</span></span>
- <span data-ttu-id="6e57c-116">`Excel.RangeCopyType.formats` 复制范围的格式设置（包括字体、颜色和其他格式），但不会复制任何值。</span><span class="sxs-lookup"><span data-stu-id="6e57c-116">`Excel.RangeCopyType.formats` copies the formatting of the range, including font, color, and other format settings, but no values.</span></span>
- <span data-ttu-id="6e57c-117">`Excel.RangeCopyType.all` (默认选项) 复制数据和格式，并保留单元格的公式（如果找到）。</span><span class="sxs-lookup"><span data-stu-id="6e57c-117">`Excel.RangeCopyType.all` (the default option) copies both data and formatting, preserving cells' formulas if found.</span></span>

<span data-ttu-id="6e57c-118">`skipBlanks` 设置是否将空白单元格复制到目标。</span><span class="sxs-lookup"><span data-stu-id="6e57c-118">`skipBlanks` sets whether blank cells are copied into the destination.</span></span> <span data-ttu-id="6e57c-119">如果为 true，`copyFrom` 将跳过源范围中的空白单元格。</span><span class="sxs-lookup"><span data-stu-id="6e57c-119">When true, `copyFrom` skips blank cells in the source range.</span></span>
<span data-ttu-id="6e57c-120">跳过的单元格不会覆盖目标范围中其对应单元格的现有数据。</span><span class="sxs-lookup"><span data-stu-id="6e57c-120">Skipped cells will not overwrite the existing data of their corresponding cells in the destination range.</span></span> <span data-ttu-id="6e57c-121">默认值为 false。</span><span class="sxs-lookup"><span data-stu-id="6e57c-121">The default is false.</span></span>

<span data-ttu-id="6e57c-122">`transpose` 确定是否将数据转置（即切换其行和列）到源位置。</span><span class="sxs-lookup"><span data-stu-id="6e57c-122">`transpose` determines whether or not the data is transposed, meaning its rows and columns are switched, into the source location.</span></span>
<span data-ttu-id="6e57c-123">转置范围沿主对角线翻转，因此，行“1”、“2”和“3”将成为列“A”、“B”和“C”。</span><span class="sxs-lookup"><span data-stu-id="6e57c-123">A transposed range is flipped along the main diagonal, so rows **1**, **2**, and **3** will become columns **A**, **B**, and **C**.</span></span>

<span data-ttu-id="6e57c-124">以下代码示例和图像在一个简单的方案中演示此行为。</span><span class="sxs-lookup"><span data-stu-id="6e57c-124">The following code sample and images demonstrate this behavior in a simple scenario.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    // copy a range, omitting the blank cells so existing data is not overwritten in those cells
    sheet.getRange("D1").copyFrom("A1:C1",
        Excel.RangeCopyType.all,
        true, // skipBlanks
        false); // transpose
    // copy a range, including the blank cells which will overwrite existing data in the target cells
    sheet.getRange("D2").copyFrom("A2:C2",
        Excel.RangeCopyType.all,
        false, // skipBlanks
        false); // transpose
    return context.sync();
}).catch(errorHandlerFunction);
```

### <a name="data-before-range-is-copied-and-pasted"></a><span data-ttu-id="6e57c-125">复制和粘贴区域之前的数据</span><span class="sxs-lookup"><span data-stu-id="6e57c-125">Data before range is copied and pasted</span></span>

![区域Excel方法运行之前的数据。](../images/excel-range-copyfrom-skipblanks-before.png)

### <a name="data-after-range-is-copied-and-pasted"></a><span data-ttu-id="6e57c-127">复制和粘贴区域之后的数据</span><span class="sxs-lookup"><span data-stu-id="6e57c-127">Data after range is copied and pasted</span></span>

![区域Excel复制方法之后的数据。](../images/excel-range-copyfrom-skipblanks-after.png)

## <a name="cut-and-paste-move-cells"></a><span data-ttu-id="6e57c-129">剪切并粘贴 (单元格) 移动</span><span class="sxs-lookup"><span data-stu-id="6e57c-129">Cut and paste (move) cells</span></span>

<span data-ttu-id="6e57c-130">[Range.moveTo](/javascript/api/excel/excel.range#moveto-destinationrange-)方法将单元格移动到工作簿中的新位置。</span><span class="sxs-lookup"><span data-stu-id="6e57c-130">The [Range.moveTo](/javascript/api/excel/excel.range#moveto-destinationrange-) method moves cells to a new location in the workbook.</span></span> <span data-ttu-id="6e57c-131">此单元格移动行为的工作方式与通过拖动区域边框或执行"[](https://support.office.com/article/Move-or-copy-cells-and-cell-contents-803d65eb-6a3e-4534-8c6f-ff12d1c4139e)剪切"和"粘贴"操作移动单元格 **时相同**。</span><span class="sxs-lookup"><span data-stu-id="6e57c-131">This cell movement behavior works the same as when cells are moved by [dragging the range border](https://support.office.com/article/Move-or-copy-cells-and-cell-contents-803d65eb-6a3e-4534-8c6f-ff12d1c4139e) or when taking the **Cut** and **Paste** actions.</span></span> <span data-ttu-id="6e57c-132">区域的格式和值都移至指定为 参数 `destinationRange` 的位置。</span><span class="sxs-lookup"><span data-stu-id="6e57c-132">Both the formatting and values of the range are moved to the location specified as the `destinationRange` parameter.</span></span>

<span data-ttu-id="6e57c-133">下面的代码示例使用 方法移动 `Range.moveTo` 区域。</span><span class="sxs-lookup"><span data-stu-id="6e57c-133">The following code sample moves a range with the `Range.moveTo` method.</span></span> <span data-ttu-id="6e57c-134">请注意，如果目标区域小于源范围，它将扩展以包含源内容。</span><span class="sxs-lookup"><span data-stu-id="6e57c-134">Note that if the destination range is smaller than the source, it will be expanded to encompass the source content.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("F1").values = [["Moved Range"]];

    // Move the cells "A1:E1" to "G1" (which fills the range "G1:K1").
    sheet.getRange("A1:E1").moveTo("G1");
    return context.sync();
});
```

## <a name="see-also"></a><span data-ttu-id="6e57c-135">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6e57c-135">See also</span></span>

- [<span data-ttu-id="6e57c-136">Excel 加载项中的 Word JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="6e57c-136">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="6e57c-137">使用 JavaScript API Excel单元格</span><span class="sxs-lookup"><span data-stu-id="6e57c-137">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="6e57c-138">使用 JavaScript API Excel重复项</span><span class="sxs-lookup"><span data-stu-id="6e57c-138">Remove duplicates using the Excel JavaScript API</span></span>](excel-add-ins-ranges-remove-duplicates.md)
- [<span data-ttu-id="6e57c-139"> 同时在 Excel 加载项中处理多个区域 </span><span class="sxs-lookup"><span data-stu-id="6e57c-139">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
