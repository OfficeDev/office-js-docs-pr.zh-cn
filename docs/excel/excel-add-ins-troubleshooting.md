---
title: Excel 外接程序疑难解答
description: 了解如何解决 Excel 外接程序中的开发错误。
ms.date: 09/08/2020
localization_priority: Normal
ms.openlocfilehash: 1bdd96772d3a221ca3a02e3d5dfcfa16561dd5f1
ms.sourcegitcommit: c6308cf245ac1bc66a876eaa0a7bb4a2492991ac
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2020
ms.locfileid: "47409380"
---
# <a name="troubleshooting-excel-add-ins"></a><span data-ttu-id="3bffa-103">Excel 外接程序疑难解答</span><span class="sxs-lookup"><span data-stu-id="3bffa-103">Troubleshooting Excel Add-ins</span></span>

<span data-ttu-id="3bffa-104">本文讨论了 Excel 特有的故障排除问题。</span><span class="sxs-lookup"><span data-stu-id="3bffa-104">This article discusses troubleshooting issues that are unique to Excel.</span></span> <span data-ttu-id="3bffa-105">请使用页面底部的反馈工具建议可添加到文章中的其他问题。</span><span class="sxs-lookup"><span data-stu-id="3bffa-105">Please use the feedback tool at the bottom of the page to suggest other issues that can be added to the article.</span></span>

## <a name="api-limitations-when-the-active-workbook-switches"></a><span data-ttu-id="3bffa-106">活动工作簿切换时的 API 限制</span><span class="sxs-lookup"><span data-stu-id="3bffa-106">API limitations when the active workbook switches</span></span>

<span data-ttu-id="3bffa-107">Excel 相关外接程序用于一次运行单个工作簿。</span><span class="sxs-lookup"><span data-stu-id="3bffa-107">Add-ins for Excel are intended to operate on a single workbook at a time.</span></span> <span data-ttu-id="3bffa-108">当运行加载项的工作簿获得焦点时，可能会出现错误。</span><span class="sxs-lookup"><span data-stu-id="3bffa-108">Errors can arise when a workbook that is separate from the one running the add-in gains focus.</span></span> <span data-ttu-id="3bffa-109">仅当焦点更改时要调用的特定方法时，才会发生这种情况。</span><span class="sxs-lookup"><span data-stu-id="3bffa-109">This only happens when particular methods are in the process of being called when the focus changes.</span></span>

<span data-ttu-id="3bffa-110">此工作簿开关会影响以下 Api：</span><span class="sxs-lookup"><span data-stu-id="3bffa-110">The following APIs are affected by this workbook switch:</span></span>

|<span data-ttu-id="3bffa-111">Excel JavaScript API</span><span class="sxs-lookup"><span data-stu-id="3bffa-111">Excel JavaScript API</span></span> | <span data-ttu-id="3bffa-112">引发的错误</span><span class="sxs-lookup"><span data-stu-id="3bffa-112">Error thrown</span></span> |
|--|--|
| `Chart.activate` | <span data-ttu-id="3bffa-113">GeneralException</span><span class="sxs-lookup"><span data-stu-id="3bffa-113">GeneralException</span></span> |
| `Range.select` | <span data-ttu-id="3bffa-114">GeneralException</span><span class="sxs-lookup"><span data-stu-id="3bffa-114">GeneralException</span></span> |
| `Table.clearFilters` | <span data-ttu-id="3bffa-115">GeneralException</span><span class="sxs-lookup"><span data-stu-id="3bffa-115">GeneralException</span></span> |
| `Workbook.getActiveCell`  | <span data-ttu-id="3bffa-116">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="3bffa-116">InvalidSelection</span></span>|
| `Workbook.getSelectedRange` | <span data-ttu-id="3bffa-117">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="3bffa-117">InvalidSelection</span></span>|
| `Workbook.getSelectedRanges`  | <span data-ttu-id="3bffa-118">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="3bffa-118">InvalidSelection</span></span>|
| `Worksheet.activate` | <span data-ttu-id="3bffa-119">GeneralException</span><span class="sxs-lookup"><span data-stu-id="3bffa-119">GeneralException</span></span> |
| `Worksheet.delete`  | <span data-ttu-id="3bffa-120">InvalidSelection</span><span class="sxs-lookup"><span data-stu-id="3bffa-120">InvalidSelection</span></span>|
| `Worksheet.gridlines` | <span data-ttu-id="3bffa-121">GeneralException</span><span class="sxs-lookup"><span data-stu-id="3bffa-121">GeneralException</span></span> |
| `Worksheet.showHeadings` | <span data-ttu-id="3bffa-122">GeneralException</span><span class="sxs-lookup"><span data-stu-id="3bffa-122">GeneralException</span></span> |
| `WorksheetCollection.add` | <span data-ttu-id="3bffa-123">GeneralException</span><span class="sxs-lookup"><span data-stu-id="3bffa-123">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeAt` | <span data-ttu-id="3bffa-124">GeneralException</span><span class="sxs-lookup"><span data-stu-id="3bffa-124">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeColumns` | <span data-ttu-id="3bffa-125">GeneralException</span><span class="sxs-lookup"><span data-stu-id="3bffa-125">GeneralException</span></span> |
| `WorksheetFreezePanes.freezeRows` | <span data-ttu-id="3bffa-126">GeneralException</span><span class="sxs-lookup"><span data-stu-id="3bffa-126">GeneralException</span></span> |
| `WorksheetFreezePanes.getLocationOrNullObject`| <span data-ttu-id="3bffa-127">GeneralException</span><span class="sxs-lookup"><span data-stu-id="3bffa-127">GeneralException</span></span> |
| `WorksheetFreezePanes.unfreeze` | <span data-ttu-id="3bffa-128">GeneralException</span><span class="sxs-lookup"><span data-stu-id="3bffa-128">GeneralException</span></span> |

> [!NOTE]
> <span data-ttu-id="3bffa-129">这仅适用于在 Windows 或 Mac 上打开的多个 Excel 工作簿。</span><span class="sxs-lookup"><span data-stu-id="3bffa-129">This only applies to multiple Excel workbooks open on Windows or Mac.</span></span>

## <a name="coauthoring"></a><span data-ttu-id="3bffa-130">共同创作</span><span class="sxs-lookup"><span data-stu-id="3bffa-130">Coauthoring</span></span>

<span data-ttu-id="3bffa-131">请参阅 [Excel 外接程序中](co-authoring-in-excel-add-ins.md) 用于共同创作环境中事件的模式的合著。</span><span class="sxs-lookup"><span data-stu-id="3bffa-131">See [Coauthoring in Excel add-ins](co-authoring-in-excel-add-ins.md) for patterns to use with events in a coauthoring environment.</span></span> <span data-ttu-id="3bffa-132">本文还讨论了使用某些 Api （例如）时的潜在合并冲突 [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-) 。</span><span class="sxs-lookup"><span data-stu-id="3bffa-132">The article also discusses potential merge conflicts when using certain APIs, such as [`TableRowCollection.add`](/javascript/api/excel/excel.tablerowcollection#add-index--values-).</span></span>

## <a name="see-also"></a><span data-ttu-id="3bffa-133">另请参阅</span><span class="sxs-lookup"><span data-stu-id="3bffa-133">See also</span></span>

- [<span data-ttu-id="3bffa-134">解决 Office 外接程序的开发错误</span><span class="sxs-lookup"><span data-stu-id="3bffa-134">Troubleshoot development errors with Office Add-ins</span></span>](../testing/troubleshoot-development-errors.md)
- [<span data-ttu-id="3bffa-135">排查 Office 加载项中的用户错误</span><span class="sxs-lookup"><span data-stu-id="3bffa-135">Troubleshoot user errors with Office Add-ins</span></span>](../testing/testing-and-troubleshooting.md)
