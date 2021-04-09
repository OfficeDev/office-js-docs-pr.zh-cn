---
title: 使用 Excel JavaScript API 读取或写入较大区域
description: 了解如何使用 Excel JavaScript API 读取或写入较大区域。
ms.date: 04/02/2021
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: b7a1e54d6b516889884f777bd256df8fb663c794
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652797"
---
# <a name="read-or-write-to-a-large-range-using-the-excel-javascript-api"></a><span data-ttu-id="6d188-103">使用 Excel JavaScript API 读取或写入较大区域</span><span class="sxs-lookup"><span data-stu-id="6d188-103">Read or write to a large range using the Excel JavaScript API</span></span>

<span data-ttu-id="6d188-104">本文介绍如何使用 Excel JavaScript API 处理对较大范围的读取和写入。</span><span class="sxs-lookup"><span data-stu-id="6d188-104">This article describes how to handle reading and writing to large ranges with the Excel JavaScript API.</span></span>

## <a name="run-separate-read-or-write-operations-for-large-ranges"></a><span data-ttu-id="6d188-105">对较大区域运行单独的读取或写入操作</span><span class="sxs-lookup"><span data-stu-id="6d188-105">Run separate read or write operations for large ranges</span></span>

<span data-ttu-id="6d188-106">如果某个区域包含大量单元格、值、数字格式或公式，则可能无法对区域运行 API 操作。</span><span class="sxs-lookup"><span data-stu-id="6d188-106">If a range contains a large number of cells, values, number formats, or formulas, it may not be possible to run API operations on that range.</span></span> <span data-ttu-id="6d188-107">API 将始终尽量尝试在区域内运行所请求的操作（即检索或写入指定的数据），但尝试对较大区域执行读取或写入操作可能会因资源利用率过高而导致 API 错误。</span><span class="sxs-lookup"><span data-stu-id="6d188-107">The API will always make a best attempt to run the requested operation on a range (i.e., to retrieve or write the specified data), but attempting to perform read or write operations for a large range may result in an API error due to excessive resource utilization.</span></span> <span data-ttu-id="6d188-108">为避免此类错误，建议为较大区域的较小子集运行单独的读取或写入操作，而不是尝试在较大区域内运行单个读取或写入操作。</span><span class="sxs-lookup"><span data-stu-id="6d188-108">To avoid such errors, we recommend that you run separate read or write operations for smaller subsets of a large range, instead of attempting to run a single read or write operation on a large range.</span></span>

<span data-ttu-id="6d188-109">有关系统限制的详细信息，请参阅 Office 加载项的资源限制和性能优化的 ["Excel 加载项"部分](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins)。</span><span class="sxs-lookup"><span data-stu-id="6d188-109">For details on the system limitations, see the "Excel add-ins" section of [Resource limits and performance optimization for Office Add-ins](../concepts/resource-limits-and-performance-optimization.md#excel-add-ins).</span></span>

### <a name="conditional-formatting-of-ranges"></a><span data-ttu-id="6d188-110">范围的条件格式</span><span class="sxs-lookup"><span data-stu-id="6d188-110">Conditional formatting of ranges</span></span>

<span data-ttu-id="6d188-111">范围可以根据条件将格式应用于个别单元格。</span><span class="sxs-lookup"><span data-stu-id="6d188-111">Ranges can have formats applied to individual cells based on conditions.</span></span> <span data-ttu-id="6d188-112">有关此操作的详细信息，请参阅[将条件格式应用于 Excel 范围](excel-add-ins-conditional-formatting.md)。</span><span class="sxs-lookup"><span data-stu-id="6d188-112">For more information about this, see [Apply conditional formatting to Excel ranges](excel-add-ins-conditional-formatting.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="6d188-113">另请参阅</span><span class="sxs-lookup"><span data-stu-id="6d188-113">See also</span></span>

- [<span data-ttu-id="6d188-114">Excel 加载项中的 Word JavaScript 对象模型</span><span class="sxs-lookup"><span data-stu-id="6d188-114">Excel JavaScript object model in Office Add-ins</span></span>](excel-add-ins-core-concepts.md)
- [<span data-ttu-id="6d188-115">使用 Excel JavaScript API 处理单元格</span><span class="sxs-lookup"><span data-stu-id="6d188-115">Work with cells using the Excel JavaScript API</span></span>](excel-add-ins-cells.md)
- [<span data-ttu-id="6d188-116">使用 Excel JavaScript API 读取或写入无限区域</span><span class="sxs-lookup"><span data-stu-id="6d188-116">Read or write to an unbounded range using the Excel JavaScript API</span></span>](excel-add-ins-ranges-unbounded.md)
- [<span data-ttu-id="6d188-117"> 同时在 Excel 加载项中处理多个区域 </span><span class="sxs-lookup"><span data-stu-id="6d188-117">Work with multiple ranges simultaneously in Excel add-ins</span></span>](excel-add-ins-multiple-ranges.md)
