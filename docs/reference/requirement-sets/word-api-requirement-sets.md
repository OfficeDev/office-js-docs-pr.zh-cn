---
title: Word JavaScript API 要求集
description: 针对 Word 内部版本的 Office 加载项要求集信息
ms.date: 07/17/2019
ms.prod: word
localization_priority: Priority
ms.openlocfilehash: 4af7a9c14489d148ffdc06a68ad6c26bf326abc5
ms.sourcegitcommit: 6d9b4820a62a914c50cef13af8b80ce626034c26
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/19/2019
ms.locfileid: "35804623"
---
# <a name="word-javascript-api-requirement-sets"></a><span data-ttu-id="3d0da-103">Word JavaScript API 要求集</span><span class="sxs-lookup"><span data-stu-id="3d0da-103">Word JavaScript API requirement sets</span></span>

<span data-ttu-id="3d0da-p101">要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="3d0da-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

## <a name="requirement-set-availability"></a><span data-ttu-id="3d0da-107">要求集可用性</span><span class="sxs-lookup"><span data-stu-id="3d0da-107">Requirement set availability</span></span>

<span data-ttu-id="3d0da-108">Word 加载项可在多个 Office 版本中运行，包括 Windows 版 Office 2016 或更高版本、Office 网页版、iPad 版 Office 和 Mac 版 Office。</span><span class="sxs-lookup"><span data-stu-id="3d0da-108">Word add-ins run across multiple versions of Office, including Office 2016 or later on Windows, Office for iPad, Office for Mac, and Office Online.</span></span> <span data-ttu-id="3d0da-109">下表列出了 Word 要求集、支持该要求集的 Office 主机应用程序，以及这些应用程序的内部版本或版本号。</span><span class="sxs-lookup"><span data-stu-id="3d0da-109">The following table lists the Word requirement sets, the Office host applications that support that requirement set, and the build or version numbers for those applications.</span></span>

> [!NOTE]
> <span data-ttu-id="3d0da-110">若要在任何编号的要求集中使用 API，你应该引用 CDN 上的**生产**库：https://appsforoffice.microsoft.com/lib/1/hosted/office.js。</span><span class="sxs-lookup"><span data-stu-id="3d0da-110">To use APIs in any of the numbered requirement sets, you should reference the **production** library on the CDN: https://appsforoffice.microsoft.com/lib/1/hosted/office.js.</span></span>
>
> <span data-ttu-id="3d0da-111">有关使用预览 API 的信息，请参阅 [Excel JavaScript 预览 API](word-preview-apis.md) 一文。</span><span class="sxs-lookup"><span data-stu-id="3d0da-111">For information about using preview APIs, see the [Excel JavaScript preview APIs](word-preview-apis.md) section within this article.</span></span>

|  <span data-ttu-id="3d0da-112">要求集</span><span class="sxs-lookup"><span data-stu-id="3d0da-112">Requirement set</span></span>  |   <span data-ttu-id="3d0da-113">Windows 版 Office\*</span><span class="sxs-lookup"><span data-stu-id="3d0da-113">Office on Windows\*</span></span><br><span data-ttu-id="3d0da-114">（连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d0da-114">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="3d0da-115">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="3d0da-115">Debug Office Add-ins on iPad and Mac</span></span><br><span data-ttu-id="3d0da-116">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d0da-116">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="3d0da-117">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="3d0da-117">Office apps on Mac</span></span><br><span data-ttu-id="3d0da-118">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3d0da-118">(connected to Office 365 subscription)</span></span>  | <span data-ttu-id="3d0da-119">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="3d0da-119">Office on the web</span></span>  |
|:-----|-----|:-----|:-----|:-----|
| [<span data-ttu-id="3d0da-120">预览</span><span class="sxs-lookup"><span data-stu-id="3d0da-120">Preview</span></span>](word-preview-apis.md) | <span data-ttu-id="3d0da-121">请使用最新的 Office 版本来试用预览 API（你可能需要加入 [Office 预览体验成员计划](https://products.office.com/office-insider)）</span><span class="sxs-lookup"><span data-stu-id="3d0da-121">Please use the latest Office version to try preview APIs (you may need to join the [Office Insider program](https://products.office.com/office-insider))</span></span> |
| [<span data-ttu-id="3d0da-122">WordApi 1.3</span><span class="sxs-lookup"><span data-stu-id="3d0da-122">WordApi 1.3</span></span>](word-api-1-3-requirement-set.md) | <span data-ttu-id="3d0da-123">版本 1612（内部版本 7668.1000）或更高版本</span><span class="sxs-lookup"><span data-stu-id="3d0da-123">Version 1612 (Build 7668.1000) or later</span></span>| <span data-ttu-id="3d0da-124">2017 年 3 月，2.22 或更高版本</span><span class="sxs-lookup"><span data-stu-id="3d0da-124">March 2017, 2.22 or later</span></span> | <span data-ttu-id="3d0da-125">2017 年 3 月，15.32 或更高版本</span><span class="sxs-lookup"><span data-stu-id="3d0da-125">March 2017, 15.32 or later</span></span>| <span data-ttu-id="3d0da-126">2017 年 3 月</span><span class="sxs-lookup"><span data-stu-id="3d0da-126">March 2017</span></span> |
| [<span data-ttu-id="3d0da-127">WordApi 1.2</span><span class="sxs-lookup"><span data-stu-id="3d0da-127">WordApi 1.2</span></span>](word-api-1-2-requirement-set.md) | <span data-ttu-id="3d0da-128">2015 年 12 月更新，版本 1601（内部版本 6568.1000）或更高版本</span><span class="sxs-lookup"><span data-stu-id="3d0da-128">December 2015 update, Version 1601 (Build 6568.1000) or later</span></span> | <span data-ttu-id="3d0da-129">2016 年 1 月，1.18 或更高版本</span><span class="sxs-lookup"><span data-stu-id="3d0da-129">January 2016, 1.18 or later</span></span> | <span data-ttu-id="3d0da-130">2016 年 1 月，15.19 或更高版本</span><span class="sxs-lookup"><span data-stu-id="3d0da-130">January 2016, 15.19 or later</span></span>| <span data-ttu-id="3d0da-131">2016 年 9 月</span><span class="sxs-lookup"><span data-stu-id="3d0da-131">September 2016</span></span> |
| [<span data-ttu-id="3d0da-132">WordApi 1.1</span><span class="sxs-lookup"><span data-stu-id="3d0da-132">WordApi 1.1</span></span>](word-api-1-1-requirement-set.md) | <span data-ttu-id="3d0da-133">版本 1509（内部版本 4266.1001）或更高版本</span><span class="sxs-lookup"><span data-stu-id="3d0da-133">Version 1509 (Build 4266.1001) or later</span></span>| <span data-ttu-id="3d0da-134">2016 年 1 月，1.18 或更高版本</span><span class="sxs-lookup"><span data-stu-id="3d0da-134">January 2016, 1.18 or later</span></span> | <span data-ttu-id="3d0da-135">2016 年 1 月，15.19 或更高版本</span><span class="sxs-lookup"><span data-stu-id="3d0da-135">January 2016, 15.19 or later</span></span>| <span data-ttu-id="3d0da-136">2016 年 9 月</span><span class="sxs-lookup"><span data-stu-id="3d0da-136">September 2016</span></span> |

> [!NOTE]
> <span data-ttu-id="3d0da-137">通过 MSI 安装的 Office 2016 的内部版本号为 16.0.4266.1001。</span><span class="sxs-lookup"><span data-stu-id="3d0da-137">The build number for Office 2016 installed via MSI is 16.0.4266.1001.</span></span> <span data-ttu-id="3d0da-138">此版本只包含 WordApi 1.1 要求集。</span><span class="sxs-lookup"><span data-stu-id="3d0da-138">This version only contains the WordApi 1.1 requirement set.</span></span>

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="3d0da-139">Office 版本和内部版本号</span><span class="sxs-lookup"><span data-stu-id="3d0da-139">Office versions and build numbers</span></span>

<span data-ttu-id="3d0da-140">有关 Office 版本和内部版本号的详细信息，请参阅：</span><span class="sxs-lookup"><span data-stu-id="3d0da-140">For more information about versions, build numbers, and Office Online Server, see:</span></span>

- [<span data-ttu-id="3d0da-141">更新频道发布的 Office 365 客户端版本号和内部版本号</span><span class="sxs-lookup"><span data-stu-id="3d0da-141">Version and build numbers of update channel releases for Office 365 clients</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="3d0da-142">使用的是哪一版 Office？</span><span class="sxs-lookup"><span data-stu-id="3d0da-142">What version of Office am I using?</span></span>](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [<span data-ttu-id="3d0da-143">在哪里可以找到 Office 365 客户端应用程序的版本号和内部版本号</span><span class="sxs-lookup"><span data-stu-id="3d0da-143">Where you can find the version and build number for an Office 365 client application</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)

## <a name="see-also"></a><span data-ttu-id="3d0da-144">另请参阅</span><span class="sxs-lookup"><span data-stu-id="3d0da-144">See also</span></span>

- [<span data-ttu-id="3d0da-145">Word JavaScript API 参考文档</span><span class="sxs-lookup"><span data-stu-id="3d0da-145">Word JavaScript API Reference Documentation</span></span>](/javascript/api/word)
- [<span data-ttu-id="3d0da-146">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="3d0da-146">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="3d0da-147">指定 Office 主机和 API 要求</span><span class="sxs-lookup"><span data-stu-id="3d0da-147">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="3d0da-148">Office 外接程序 XML 清单</span><span class="sxs-lookup"><span data-stu-id="3d0da-148">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
