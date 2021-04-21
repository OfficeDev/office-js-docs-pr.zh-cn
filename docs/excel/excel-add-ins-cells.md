---
title: 使用 Excel JavaScript API 处理单元格。
description: 了解单元格的 Excel JavaScript API 定义，并了解如何使用单元格。
ms.date: 04/16/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: ad8ca985b6bbdcf19920c36c371e690f61639f16
ms.sourcegitcommit: da8ad214406f2e1cd80982af8a13090e76187dbd
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/21/2021
ms.locfileid: "51917098"
---
# <a name="work-with-cells-using-the-excel-javascript-api"></a><span data-ttu-id="d0227-103">使用 Excel JavaScript API 处理单元格</span><span class="sxs-lookup"><span data-stu-id="d0227-103">Work with cells using the Excel JavaScript API</span></span>

<span data-ttu-id="d0227-104">The Excel JavaScript API 没有“Cell”对象或类。</span><span class="sxs-lookup"><span data-stu-id="d0227-104">The Excel JavaScript API doesn't have a "Cell" object or class.</span></span> <span data-ttu-id="d0227-105">相反，所有 Excel 单元格都是 `Range` 对象。</span><span class="sxs-lookup"><span data-stu-id="d0227-105">Instead, all Excel cells are `Range` objects.</span></span> <span data-ttu-id="d0227-106">Excel UI 中的单个单元格转换为 Excel JavaScript API 中包含一个单元格的 `Range` 对象。</span><span class="sxs-lookup"><span data-stu-id="d0227-106">An individual cell in the Excel UI translates to a `Range` object with one cell in the Excel JavaScript API.</span></span>

<span data-ttu-id="d0227-107">对象 `Range` 还可以包含多个连续单元格。</span><span class="sxs-lookup"><span data-stu-id="d0227-107">A `Range` object can also contain multiple, contiguous cells.</span></span> <span data-ttu-id="d0227-108">连续单元格形成一个不间断的矩形 (包括单个行或) 。</span><span class="sxs-lookup"><span data-stu-id="d0227-108">Contiguous cells form an unbroken rectangle (including single rows or columns).</span></span> <span data-ttu-id="d0227-109">若要了解如何处理不连续的单元格，请参阅使用 [RangeAreas](#work-with-discontiguous-cells-using-the-rangeareas-object)对象处理不连续的单元格。</span><span class="sxs-lookup"><span data-stu-id="d0227-109">To learn about working with cells that are not contiguous, see [Work with discontiguous cells using the RangeAreas object](#work-with-discontiguous-cells-using-the-rangeareas-object).</span></span>

<span data-ttu-id="d0227-110">有关对象支持的属性和方法的完整列表，请参阅 `Range` Range Object [ (JavaScript API for Excel) 。 ](/javascript/api/excel/excel.range)</span><span class="sxs-lookup"><span data-stu-id="d0227-110">For the complete list of properties and methods that the `Range` object supports, see [Range Object (JavaScript API for Excel)](/javascript/api/excel/excel.range).</span></span>

## <a name="work-with-discontiguous-cells-using-the-rangeareas-object"></a><span data-ttu-id="d0227-111">使用 RangeAreas 对象处理不连续单元格</span><span class="sxs-lookup"><span data-stu-id="d0227-111">Work with discontiguous cells using the RangeAreas object</span></span>

<span data-ttu-id="d0227-112">[RangeAreas](/javascript/api/excel/excel.rangeareas)对象允许您的外接程序一次对多个区域执行操作。</span><span class="sxs-lookup"><span data-stu-id="d0227-112">The [RangeAreas](/javascript/api/excel/excel.rangeareas) object lets your add-in perform operations on multiple ranges at once.</span></span> <span data-ttu-id="d0227-113">这些区域可能是连续的，但不必是。</span><span class="sxs-lookup"><span data-stu-id="d0227-113">These ranges may be contiguous, but they don't have to be.</span></span> <span data-ttu-id="d0227-114">`RangeAreas` 将进一步在[同时在 Excel 加载项中处理多个区域](excel-add-ins-multiple-ranges.md)一文中进行讨论。</span><span class="sxs-lookup"><span data-stu-id="d0227-114">`RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="d0227-115">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d0227-115">See also</span></span>

- [<span data-ttu-id="d0227-116">Excel 加载项中的 Word JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="d0227-116">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="d0227-117">使用 Excel JavaScript API 获取区域</span><span class="sxs-lookup"><span data-stu-id="d0227-117">Get a range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-get.md)
- [<span data-ttu-id="d0227-118"> 同时在 Excel 加载项中处理多个区域 </span><span class="sxs-lookup"><span data-stu-id="d0227-118">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
