---
title: 使用 JavaScript API Excel重复项
description: 了解如何使用 JavaScript API Excel删除重复项。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 859214d36bdf66a284304ba1d5f7f2d642b718cb
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/23/2021
ms.locfileid: "53075766"
---
# <a name="remove-duplicates-using-the-excel-javascript-api"></a><span data-ttu-id="87556-103">使用 JavaScript API Excel重复项</span><span class="sxs-lookup"><span data-stu-id="87556-103">Remove duplicates using the Excel JavaScript API</span></span>

<span data-ttu-id="87556-104">本文提供了一个代码示例，该示例使用 JavaScript API 删除Excel条目。</span><span class="sxs-lookup"><span data-stu-id="87556-104">This article provides a code sample that removes duplicate entries in a range using the Excel JavaScript API.</span></span> <span data-ttu-id="87556-105">有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel。Range 类](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="87556-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="remove-rows-with-duplicate-entries"></a><span data-ttu-id="87556-106">删除条目重复的行</span><span class="sxs-lookup"><span data-stu-id="87556-106">Remove rows with duplicate entries</span></span>

<span data-ttu-id="87556-107">[Range.removeDuplicates](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-)方法删除指定列中包含重复条目的行。</span><span class="sxs-lookup"><span data-stu-id="87556-107">The [Range.removeDuplicates](/javascript/api/excel/excel.range#removeduplicates-columns--includesheader-) method removes rows with duplicate entries in the specified columns.</span></span> <span data-ttu-id="87556-108">方法将浏览区域的每一行，从最低值索引到最高值索引，从 (到底部) 。</span><span class="sxs-lookup"><span data-stu-id="87556-108">The method goes through each row in the range from the lowest-valued index to the highest-valued index in the range (from top to bottom).</span></span> <span data-ttu-id="87556-109">如果指定列中的值之前显示在区域中，则会删除该行。</span><span class="sxs-lookup"><span data-stu-id="87556-109">A row is deleted if a value in its specified column or columns appeared earlier in the range.</span></span> <span data-ttu-id="87556-110">在区域内位于已删除行下方的行将上移。</span><span class="sxs-lookup"><span data-stu-id="87556-110">Rows in the range below the deleted row are shifted up.</span></span> <span data-ttu-id="87556-111">`removeDuplicates` 不影响该区域外的单元格位置。</span><span class="sxs-lookup"><span data-stu-id="87556-111">`removeDuplicates` does not affect the position of cells outside of the range.</span></span>

<span data-ttu-id="87556-112">`removeDuplicates` 使用 `number[]` 来表示已执行重复项检查的列索引。</span><span class="sxs-lookup"><span data-stu-id="87556-112">`removeDuplicates` takes in a `number[]` representing the column indices which are checked for duplicates.</span></span> <span data-ttu-id="87556-113">此数组从零开始并且与区域而不是与工作表相关。</span><span class="sxs-lookup"><span data-stu-id="87556-113">This array is zero-based and relative to the range, not the worksheet.</span></span> <span data-ttu-id="87556-114">该方法还采用一个布尔参数，该参数指定第一行是否是标题。</span><span class="sxs-lookup"><span data-stu-id="87556-114">The method also takes in a boolean parameter that specifies whether the first row is a header.</span></span> <span data-ttu-id="87556-115">如果为 **true**，则在考虑重复项时将忽略顶行。</span><span class="sxs-lookup"><span data-stu-id="87556-115">When **true**, the top row is ignored when considering duplicates.</span></span> <span data-ttu-id="87556-116">该方法返回一个对象，该对象指定删除的行数和剩余 `removeDuplicates` `RemoveDuplicatesResult` 的唯一行数。</span><span class="sxs-lookup"><span data-stu-id="87556-116">The `removeDuplicates` method returns a `RemoveDuplicatesResult` object that specifies the number of rows removed and the number of unique rows remaining.</span></span>

<span data-ttu-id="87556-117">使用区域的方法 `removeDuplicates` 时，请记住以下事项：</span><span class="sxs-lookup"><span data-stu-id="87556-117">When using a range's `removeDuplicates` method, keep the following in mind:</span></span>

- <span data-ttu-id="87556-118">`removeDuplicates` 会考虑单元格值，而不是函数结果。</span><span class="sxs-lookup"><span data-stu-id="87556-118">`removeDuplicates` considers cell values, not function results.</span></span> <span data-ttu-id="87556-119">如果两个不同的函数具有相同的求值结果，则不会将单元格值视为重复项。</span><span class="sxs-lookup"><span data-stu-id="87556-119">If two different functions evaluate to the same result, the cell values are not considered duplicates.</span></span>
- <span data-ttu-id="87556-120">`removeDuplicates` 不会忽略空单元格。</span><span class="sxs-lookup"><span data-stu-id="87556-120">Empty cells are not ignored by `removeDuplicates`.</span></span> <span data-ttu-id="87556-121">空单元格的值与任何其他值具有相同的处理方式。</span><span class="sxs-lookup"><span data-stu-id="87556-121">The value of an empty cell is treated like any other value.</span></span> <span data-ttu-id="87556-122">这意味着区域内所含的空行将包含在 `RemoveDuplicatesResult` 中。</span><span class="sxs-lookup"><span data-stu-id="87556-122">This means empty rows contained within in the range will be included in the `RemoveDuplicatesResult`.</span></span>

<span data-ttu-id="87556-123">下面的代码示例演示删除第一列中具有重复值的条目。</span><span class="sxs-lookup"><span data-stu-id="87556-123">The following code sample shows the removal of entries with duplicate values in the first column.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");
    var range = sheet.getRange("B2:D11");

    var deleteResult = range.removeDuplicates([0],true);
    deleteResult.load();

    return context.sync().then(function () {
        console.log(deleteResult.removed + " entries with duplicate names removed.");
        console.log(deleteResult.uniqueRemaining + " entries with unique names remain in the range.");
    });
}).catch(errorHandlerFunction);
```

### <a name="data-before-duplicate-entries-are-removed"></a><span data-ttu-id="87556-124">删除重复条目之前的数据</span><span class="sxs-lookup"><span data-stu-id="87556-124">Data before duplicate entries are removed</span></span>

![区域Excel重复项方法之前的数据。](../images/excel-ranges-remove-duplicates-before.png)

### <a name="data-after-duplicate-entries-are-removed"></a><span data-ttu-id="87556-126">删除重复条目后的数据</span><span class="sxs-lookup"><span data-stu-id="87556-126">Data after duplicate entries are removed</span></span>

![区域Excel重复项方法运行后的数据。](../images/excel-ranges-remove-duplicates-after.png)

## <a name="see-also"></a><span data-ttu-id="87556-128">另请参阅</span><span class="sxs-lookup"><span data-stu-id="87556-128">See also</span></span>

- [<span data-ttu-id="87556-129">Excel 加载项中的 Word JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="87556-129">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="87556-130">使用 JavaScript API Excel单元格</span><span class="sxs-lookup"><span data-stu-id="87556-130">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="87556-131">使用 JavaScript API Excel、复制和粘贴区域</span><span class="sxs-lookup"><span data-stu-id="87556-131">Cut, copy, and paste ranges using the Excel JavaScript API</span></span>](excel-add-ins-ranges-cut-copy-paste.md)
- [<span data-ttu-id="87556-132"> 同时在 Excel 加载项中处理多个区域 </span><span class="sxs-lookup"><span data-stu-id="87556-132">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
