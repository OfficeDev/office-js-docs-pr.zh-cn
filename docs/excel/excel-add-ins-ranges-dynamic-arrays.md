---
title: 使用 Excel JavaScript API 处理动态数组和范围溢出
description: 了解如何使用 Excel JavaScript API 处理动态数组和范围溢出。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: c224fc336791440911519a6d24aee6c208d90c9e
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652804"
---
# <a name="handle-dynamic-arrays-and-spilling-using-the-excel-javascript-api"></a><span data-ttu-id="105ae-103">使用 Excel JavaScript API 处理动态数组和溢出</span><span class="sxs-lookup"><span data-stu-id="105ae-103">Handle dynamic arrays and spilling using the Excel JavaScript API</span></span>

<span data-ttu-id="105ae-104">本文提供了一个代码示例，该示例使用 Excel JavaScript API 处理动态数组和范围溢出。</span><span class="sxs-lookup"><span data-stu-id="105ae-104">This article provides a code sample that handles dynamic arrays and range spilling using the Excel JavaScript API.</span></span> <span data-ttu-id="105ae-105">有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel.Range 类](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="105ae-105">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="dynamic-arrays"></a><span data-ttu-id="105ae-106">动态数组</span><span class="sxs-lookup"><span data-stu-id="105ae-106">Dynamic arrays</span></span>

<span data-ttu-id="105ae-107">一些 Excel 公式返回 [动态数组](https://support.microsoft.com/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531)。</span><span class="sxs-lookup"><span data-stu-id="105ae-107">Some Excel formulas return [Dynamic arrays](https://support.microsoft.com/office/dynamic-array-formulas-and-spilled-array-behavior-205c6b06-03ba-4151-89a1-87a7eb36e531).</span></span> <span data-ttu-id="105ae-108">这些填充公式原始单元格之外的多个单元格的值。</span><span class="sxs-lookup"><span data-stu-id="105ae-108">These fill the values of multiple cells outside of the formula's original cell.</span></span> <span data-ttu-id="105ae-109">此值溢出称为"溢出"。</span><span class="sxs-lookup"><span data-stu-id="105ae-109">This value overflow is referred to as a "spill".</span></span> <span data-ttu-id="105ae-110">外接程序可以使用 [Range.getSpillingToRange](/javascript/api/excel/excel.range#getspillingtorange--) 方法查找用于溢出的范围。</span><span class="sxs-lookup"><span data-stu-id="105ae-110">Your add-in can find the range used for a spill with the [Range.getSpillingToRange](/javascript/api/excel/excel.range#getspillingtorange--) method.</span></span> <span data-ttu-id="105ae-111">还有 [\*OrNullObject 版本](..//develop/application-specific-api-model.md#ornullobject-methods-and-properties) `Range.getSpillingToRangeOrNullObject` 。</span><span class="sxs-lookup"><span data-stu-id="105ae-111">There is also a [\*OrNullObject version](..//develop/application-specific-api-model.md#ornullobject-methods-and-properties), `Range.getSpillingToRangeOrNullObject`.</span></span>

<span data-ttu-id="105ae-112">以下示例显示一个基本公式，该公式将区域的内容复制到单元格中，该公式会溢出到相邻的单元格中。</span><span class="sxs-lookup"><span data-stu-id="105ae-112">The following sample shows a basic formula that copies the contents of a range into a cell, which spills into neighboring cells.</span></span> <span data-ttu-id="105ae-113">然后，外接程序记录包含溢出的范围。</span><span class="sxs-lookup"><span data-stu-id="105ae-113">The add-in then logs the range that contains the spill.</span></span>

```js
Excel.run(function (context) {
    var sheet = context.workbook.worksheets.getItem("Sample");

    // Set G4 to a formula that returns a dynamic array.
    var targetCell = sheet.getRange("G4");
    targetCell.formulas = [["=A4:D4"]];

    // Get the address of the cells that the dynamic array spilled into.
    var spillRange = targetCell.getSpillingToRange();
    spillRange.load("address");

    // Sync and log the spilled-to range.
    return context.sync().then(function () {
        // This will log the range as "G4:J4".
        console.log(`Copying the table headers spilled into ${spillRange.address}.`);
    });
}).catch(errorHandlerFunction);
```

## <a name="range-spilling"></a><span data-ttu-id="105ae-114">区域溢出</span><span class="sxs-lookup"><span data-stu-id="105ae-114">Range spilling</span></span>

<span data-ttu-id="105ae-115">使用 [Range.getSpillParent](/javascript/api/excel/excel.range#getspillparent--) 方法查找负责溢出到给定单元格的单元格。</span><span class="sxs-lookup"><span data-stu-id="105ae-115">Find the cell responsible for spilling into a given cell by using the [Range.getSpillParent](/javascript/api/excel/excel.range#getspillparent--) method.</span></span> <span data-ttu-id="105ae-116">请注意， `getSpillParent` 仅在 range 对象是单个单元格时有效。</span><span class="sxs-lookup"><span data-stu-id="105ae-116">Note that `getSpillParent` only works when the range object is a single cell.</span></span> <span data-ttu-id="105ae-117">对 `getSpillParent` 具有多个单元格的范围调用将导致在返回 (返回空区域时 `Range.getSpillParentOrNullObject` 引发错误) 。</span><span class="sxs-lookup"><span data-stu-id="105ae-117">Calling `getSpillParent` on a range with multiple cells will result in an error being thrown (or a null range being returned for `Range.getSpillParentOrNullObject`).</span></span>

## <a name="see-also"></a><span data-ttu-id="105ae-118">另请参阅</span><span class="sxs-lookup"><span data-stu-id="105ae-118">See also</span></span>

- [<span data-ttu-id="105ae-119">Excel 加载项中的 Word JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="105ae-119">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="105ae-120">使用 Excel JavaScript API 处理单元格</span><span class="sxs-lookup"><span data-stu-id="105ae-120">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="105ae-121"> 同时在 Excel 加载项中处理多个区域 </span><span class="sxs-lookup"><span data-stu-id="105ae-121">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
