---
title: Word JavaScript API 要求集
description: 针对 Word 内部版本的 Office 加载项要求集信息。
ms.date: 04/16/2020
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: bffd78455cd6d87a1323c4133ce16f9723e37a4c
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611281"
---
# <a name="word-javascript-api-requirement-sets"></a><span data-ttu-id="601e9-103">Word JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="601e9-103">Word JavaScript API requirement sets</span></span>

<span data-ttu-id="601e9-p101">要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="601e9-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

## <a name="requirement-set-availability"></a><span data-ttu-id="601e9-107">要求集可用性</span><span class="sxs-lookup"><span data-stu-id="601e9-107">Requirement set availability</span></span>

<span data-ttu-id="601e9-108">Word 加载项可在多个 Office 版本中运行，包括 Windows 版 Office 2016 或更高版本、Office 网页版、iPad 版 Office 和 Mac 版 Office。</span><span class="sxs-lookup"><span data-stu-id="601e9-108">Word add-ins run across multiple versions of Office, including Office 2016 or later on Windows, and Office on the web, iPad, and Mac.</span></span> <span data-ttu-id="601e9-109">下表列出了 Word 要求集、支持该要求集的 Office 主机应用程序，以及这些应用程序的内部版本或版本号。</span><span class="sxs-lookup"><span data-stu-id="601e9-109">The following table lists the Word requirement sets, the Office host applications that support that requirement set, and the build or version numbers for those applications.</span></span>

> [!NOTE]
> <span data-ttu-id="601e9-110">若要在任何编号的要求集中使用 API，你应该引用 CDN 上的**生产**库：https://appsforoffice.microsoft.com/lib/1/hosted/office.js。</span><span class="sxs-lookup"><span data-stu-id="601e9-110">To use APIs in any of the numbered requirement sets, you should reference the **production** library on the CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js.</span></span>
>
> <span data-ttu-id="601e9-111">有关使用预览 API 的信息，请参阅 [Excel JavaScript 预览 API](word-preview-apis.md) 一文。</span><span class="sxs-lookup"><span data-stu-id="601e9-111">For information about using preview APIs, see the [Excel JavaScript preview APIs](word-preview-apis.md) article.</span></span>

|  <span data-ttu-id="601e9-112">要求集</span><span class="sxs-lookup"><span data-stu-id="601e9-112">Requirement set</span></span>  |   <span data-ttu-id="601e9-113">Windows 版 Office\*</span><span class="sxs-lookup"><span data-stu-id="601e9-113">Office on Windows\*</span></span><br><span data-ttu-id="601e9-114">（连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="601e9-114">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="601e9-115">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="601e9-115">Office on iPad</span></span><br><span data-ttu-id="601e9-116">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="601e9-116">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="601e9-117">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="601e9-117">Office on Mac</span></span><br><span data-ttu-id="601e9-118">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="601e9-118">(connected to Office 365 subscription)</span></span>  | <span data-ttu-id="601e9-119">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="601e9-119">Office on the web</span></span>  |
|:-----|-----|:-----|:-----|:-----|
| [<span data-ttu-id="601e9-120">预览</span><span class="sxs-lookup"><span data-stu-id="601e9-120">Preview</span></span>](word-preview-apis.md) | <span data-ttu-id="601e9-121">请使用最新的 Office 版本来试用预览 API（你可能需要加入 [Office 预览体验成员计划](https://insider.office.com)）</span><span class="sxs-lookup"><span data-stu-id="601e9-121">Please use the latest Office version to try preview APIs (you may need to join the [Office Insider program](https://insider.office.com))</span></span> |
| [<span data-ttu-id="601e9-122">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="601e9-122">WordApi 1.3</span></span>](word-api-1-3-requirement-set.md) | <span data-ttu-id="601e9-123">版本 1612（内部版本 7668.1000）或更高版本</span><span class="sxs-lookup"><span data-stu-id="601e9-123">Version 1612 (Build 7668.1000) or later</span></span>| <span data-ttu-id="601e9-124">2017 年 3 月，2.22 或更高版本</span><span class="sxs-lookup"><span data-stu-id="601e9-124">March 2017, 2.22 or later</span></span> | <span data-ttu-id="601e9-125">2017 年 3 月，15.32 或更高版本</span><span class="sxs-lookup"><span data-stu-id="601e9-125">March 2017, 15.32 or later</span></span>| <span data-ttu-id="601e9-126">2017 年 3 月</span><span class="sxs-lookup"><span data-stu-id="601e9-126">March 2017</span></span> |
| [<span data-ttu-id="601e9-127">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="601e9-127">WordApi 1.2</span></span>](word-api-1-2-requirement-set.md) | <span data-ttu-id="601e9-128">2015 年 12 月更新，版本 1601（内部版本 6568.1000）或更高版本</span><span class="sxs-lookup"><span data-stu-id="601e9-128">December 2015 update, Version 1601 (Build 6568.1000) or later</span></span> | <span data-ttu-id="601e9-129">2016 年 1 月，1.18 或更高版本</span><span class="sxs-lookup"><span data-stu-id="601e9-129">January 2016, 1.18 or later</span></span> | <span data-ttu-id="601e9-130">2016 年 1 月，15.19 或更高版本</span><span class="sxs-lookup"><span data-stu-id="601e9-130">January 2016, 15.19 or later</span></span>| <span data-ttu-id="601e9-131">2016 年 9 月</span><span class="sxs-lookup"><span data-stu-id="601e9-131">September 2016</span></span> |
| [<span data-ttu-id="601e9-132">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="601e9-132">WordApi 1.1</span></span>](word-api-1-1-requirement-set.md) | <span data-ttu-id="601e9-133">版本 1509（内部版本 4266.1001）或更高版本</span><span class="sxs-lookup"><span data-stu-id="601e9-133">Version 1509 (Build 4266.1001) or later</span></span>| <span data-ttu-id="601e9-134">2016 年 1 月，1.18 或更高版本</span><span class="sxs-lookup"><span data-stu-id="601e9-134">January 2016, 1.18 or later</span></span> | <span data-ttu-id="601e9-135">2016 年 1 月，15.19 或更高版本</span><span class="sxs-lookup"><span data-stu-id="601e9-135">January 2016, 15.19 or later</span></span>| <span data-ttu-id="601e9-136">2016 年 9 月</span><span class="sxs-lookup"><span data-stu-id="601e9-136">September 2016</span></span> |

> [!NOTE]
> <span data-ttu-id="601e9-137">永久版本的 Office 支持要求集如下：</span><span class="sxs-lookup"><span data-stu-id="601e9-137">Perpetual versions of Office support requirement sets as follows:</span></span>
>
> - <span data-ttu-id="601e9-138">Office 2019 支持 ExcelApi 1.3 及更低版本。</span><span class="sxs-lookup"><span data-stu-id="601e9-138">Office 2019 supports WordApi 1.3 and earlier.</span></span>
> - <span data-ttu-id="601e9-139">Office 2016 仅支持 ExcelApi 1.1 要求集。</span><span class="sxs-lookup"><span data-stu-id="601e9-139">Office 2016 only supports the WordApi 1.1 requirement set.</span></span>

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="601e9-140">Office 版本和内部版本号</span><span class="sxs-lookup"><span data-stu-id="601e9-140">Office versions and build numbers</span></span>

<span data-ttu-id="601e9-141">有关 Office 版本和内部版本号的详细信息，请参阅：</span><span class="sxs-lookup"><span data-stu-id="601e9-141">For more information about Office versions and build numbers, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]

## <a name="see-also"></a><span data-ttu-id="601e9-142">另请参阅</span><span class="sxs-lookup"><span data-stu-id="601e9-142">See also</span></span>

- [<span data-ttu-id="601e9-143">Word JavaScript API 参考文档</span><span class="sxs-lookup"><span data-stu-id="601e9-143">Word JavaScript API Reference Documentation</span></span>](/javascript/api/word)
- [<span data-ttu-id="601e9-144">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="601e9-144">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="601e9-145">指定 Office 主机和 API 要求</span><span class="sxs-lookup"><span data-stu-id="601e9-145">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="601e9-146">Office 外接程序 XML 清单</span><span class="sxs-lookup"><span data-stu-id="601e9-146">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
