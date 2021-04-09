---
title: 使用 Excel JavaScript API 处理单元格。
description: 了解单元格的 Excel JavaScript API 定义，并了解如何使用单元格。
ms.date: 04/07/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 5fcfeeef52f17c22d13ed3c1a10851f1d8e69204
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652874"
---
# <a name="work-with-cells-using-the-excel-javascript-api"></a><span data-ttu-id="d248b-103">使用 Excel JavaScript API 处理单元格</span><span class="sxs-lookup"><span data-stu-id="d248b-103">Work with cells using the Excel JavaScript API</span></span>

<span data-ttu-id="d248b-104">Excel JavaScript API 没有"Cell"对象或类。</span><span class="sxs-lookup"><span data-stu-id="d248b-104">The Excel JavaScript API doesn't have a "Cell" object or class.</span></span> <span data-ttu-id="d248b-105">相反，所有 Excel 单元格都是 `Range` 对象。</span><span class="sxs-lookup"><span data-stu-id="d248b-105">Instead, all Excel cells are `Range` objects.</span></span> <span data-ttu-id="d248b-106">Excel UI 中的单个单元格转换为 Excel JavaScript API 中具有一个单元格 `Range` 的对象。</span><span class="sxs-lookup"><span data-stu-id="d248b-106">An individual cell in the Excel UI translates to a `Range` object with one cell in the Excel JavaScript API.</span></span>

<span data-ttu-id="d248b-107">对象 `Range` 还可以包含多个连续单元格。</span><span class="sxs-lookup"><span data-stu-id="d248b-107">A `Range` object can also contain multiple, contiguous cells.</span></span> <span data-ttu-id="d248b-108">连续单元格形成一个不间断的矩形 (包括单个行或) 。</span><span class="sxs-lookup"><span data-stu-id="d248b-108">Contiguous cells form an unbroken rectangle (including single rows or columns).</span></span> <span data-ttu-id="d248b-109">若要了解如何处理不连续的单元格，请参阅使用 [RangeAreas](#work-with-discontiguous-cells-using-the-rangeareas-object)对象处理不连续的单元格。</span><span class="sxs-lookup"><span data-stu-id="d248b-109">To learn about working with cells that are not contiguous, see [Work with discontiguous cells using the RangeAreas object](#work-with-discontiguous-cells-using-the-rangeareas-object).</span></span>

<span data-ttu-id="d248b-110">有关对象支持的属性和方法的完整列表，请参阅 `Range` [Excel.Range 类](/javascript/api/excel/excel.range)。</span><span class="sxs-lookup"><span data-stu-id="d248b-110">For the complete list of properties and methods that the `Range` object supports, see [Excel.Range class](/javascript/api/excel/excel.range).</span></span>

## <a name="excel-javascript-apis-that-mention-cells"></a><span data-ttu-id="d248b-111">提及单元格的 Excel JavaScript API</span><span class="sxs-lookup"><span data-stu-id="d248b-111">Excel JavaScript APIs that mention cells</span></span>

<span data-ttu-id="d248b-112">即使 Excel JavaScript API 没有"Cell"对象或类，许多 API 名称也提及单元格。</span><span class="sxs-lookup"><span data-stu-id="d248b-112">Even though the Excel JavaScript API doesn't have a "Cell" object or class, a number of API names mention cells.</span></span> <span data-ttu-id="d248b-113">这些 API 控制单元格属性，如颜色、文本格式和字体。</span><span class="sxs-lookup"><span data-stu-id="d248b-113">These APIs control cell properties like color, text formatting, and font.</span></span>

<span data-ttu-id="d248b-114">以下 Excel JavaScript API 列表引用单元格。</span><span class="sxs-lookup"><span data-stu-id="d248b-114">The following list of Excel JavaScript APIs refer to cells.</span></span>

- [<span data-ttu-id="d248b-115">CellBorder</span><span class="sxs-lookup"><span data-stu-id="d248b-115">CellBorder</span></span>](/javascript/api/excel/excel.cellborder)
- [<span data-ttu-id="d248b-116">CellBorderCollection</span><span class="sxs-lookup"><span data-stu-id="d248b-116">CellBorderCollection</span></span>](/javascript/api/excel/excel.cellbordercollection)
- [<span data-ttu-id="d248b-117">CellProperties</span><span class="sxs-lookup"><span data-stu-id="d248b-117">CellProperties</span></span>](/javascript/api/excel/excel.cellproperties)
- [<span data-ttu-id="d248b-118">CellPropertiesFill</span><span class="sxs-lookup"><span data-stu-id="d248b-118">CellPropertiesFill</span></span>](/javascript/api/excel/excel.cellpropertiesfill)
- [<span data-ttu-id="d248b-119">CellPropertiesFont</span><span class="sxs-lookup"><span data-stu-id="d248b-119">CellPropertiesFont</span></span>](/javascript/api/excel/excel.cellpropertiesfont)
- [<span data-ttu-id="d248b-120">CellPropertiesFormat</span><span class="sxs-lookup"><span data-stu-id="d248b-120">CellPropertiesFormat</span></span>](/javascript/api/excel/excel.cellpropertiesformat)
- [<span data-ttu-id="d248b-121">CellPropertiesProtection</span><span class="sxs-lookup"><span data-stu-id="d248b-121">CellPropertiesProtection</span></span>](/javascript/api/excel/excel.cellpropertiesprotection)
- [<span data-ttu-id="d248b-122">CellValueConditionalFormat</span><span class="sxs-lookup"><span data-stu-id="d248b-122">CellValueConditionalFormat</span></span>](/javascript/api/excel/excel.cellvalueconditionalformat)
- [<span data-ttu-id="d248b-123">ConditionalCellValueRule</span><span class="sxs-lookup"><span data-stu-id="d248b-123">ConditionalCellValueRule</span></span>](/javascript/api/excel/excel.conditionalcellvaluerule)
- [<span data-ttu-id="d248b-124">SettableCellProperties</span><span class="sxs-lookup"><span data-stu-id="d248b-124">SettableCellProperties</span></span>](/javascript/api/excel/excel.settablecellproperties)

## <a name="work-with-discontiguous-cells-using-the-rangeareas-object"></a><span data-ttu-id="d248b-125">使用 RangeAreas 对象处理不连续单元格</span><span class="sxs-lookup"><span data-stu-id="d248b-125">Work with discontiguous cells using the RangeAreas object</span></span>

<span data-ttu-id="d248b-126">[RangeAreas](/javascript/api/excel/excel.rangeareas)对象允许您的外接程序一次对多个区域执行操作。</span><span class="sxs-lookup"><span data-stu-id="d248b-126">The [RangeAreas](/javascript/api/excel/excel.rangeareas) object lets your add-in perform operations on multiple ranges at once.</span></span> <span data-ttu-id="d248b-127">这些区域可能是连续的，但不必是。</span><span class="sxs-lookup"><span data-stu-id="d248b-127">These ranges may be contiguous, but they don't have to be.</span></span> <span data-ttu-id="d248b-128">`RangeAreas` 将进一步在[同时在 Excel 加载项中处理多个区域](excel-add-ins-multiple-ranges.md)一文中进行讨论。</span><span class="sxs-lookup"><span data-stu-id="d248b-128">`RangeAreas` are further discussed in the article [Work with multiple ranges simultaneously in Excel add-ins](excel-add-ins-multiple-ranges.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="d248b-129">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d248b-129">See also</span></span>

- [<span data-ttu-id="d248b-130">Excel 加载项中的 Word JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="d248b-130">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="d248b-131">使用 Excel JavaScript API 获取区域</span><span class="sxs-lookup"><span data-stu-id="d248b-131">Get a range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-get.md)
- [<span data-ttu-id="d248b-132"> 同时在 Excel 加载项中处理多个区域 </span><span class="sxs-lookup"><span data-stu-id="d248b-132">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
