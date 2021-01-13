---
title: 功能区 API 要求集
description: 指定支持动态功能区 API 的 Office 平台和内部版本。
ms.date: 11/07/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 91c909755779d122fba8d77dc246784f6a0dd1a3
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839983"
---
# <a name="ribbon-api-requirement-sets"></a><span data-ttu-id="4cb81-103">功能区 API 要求集</span><span class="sxs-lookup"><span data-stu-id="4cb81-103">Ribbon API requirement sets</span></span>

<span data-ttu-id="4cb81-p101">要求集是指已命名的 API 成员组。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="4cb81-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="4cb81-107">功能区 API 集支持以编程方式控制自定义外接程序命令 (，即自定义功能区按钮和菜单项) 和禁用。</span><span class="sxs-lookup"><span data-stu-id="4cb81-107">The Ribbon API set supports programmatic control of when custom Add-in Commands (that is, custom ribbon buttons and menu items) are enabled and disabled.</span></span>

<span data-ttu-id="4cb81-108">Office 外接程序在多个 Office 版本中运行。</span><span class="sxs-lookup"><span data-stu-id="4cb81-108">Office Add-ins run across multiple versions of Office.</span></span> <span data-ttu-id="4cb81-109">下表列出了功能区 API 要求集、支持该要求集的 Office 客户端应用程序，以及 Office 应用程序内部版本或版本号。</span><span class="sxs-lookup"><span data-stu-id="4cb81-109">The following table lists the Ribbon API requirement sets, the Office client applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="4cb81-110">要求集</span><span class="sxs-lookup"><span data-stu-id="4cb81-110">Requirement set</span></span>  | <span data-ttu-id="4cb81-111">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="4cb81-111">Office 2013 on Windows</span></span><br><span data-ttu-id="4cb81-112">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="4cb81-112">(one-time purchase)</span></span> | <span data-ttu-id="4cb81-113">Windows 版 Office 2016 或更高版本</span><span class="sxs-lookup"><span data-stu-id="4cb81-113">Office 2016 or later on Windows</span></span><br><span data-ttu-id="4cb81-114">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="4cb81-114">(one-time purchase)</span></span>   | <span data-ttu-id="4cb81-115">Windows 版 Office\*</span><span class="sxs-lookup"><span data-stu-id="4cb81-115">Office on Windows\*</span></span><br><span data-ttu-id="4cb81-116">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="4cb81-116">(connected to a Microsoft 365 subscription)</span></span> |  <span data-ttu-id="4cb81-117">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="4cb81-117">Office on iPad</span></span><br><span data-ttu-id="4cb81-118">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="4cb81-118">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="4cb81-119">Mac 版 Office\*</span><span class="sxs-lookup"><span data-stu-id="4cb81-119">Office on Mac\*</span></span><br><span data-ttu-id="4cb81-120">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="4cb81-120">(connected to a Microsoft 365 subscription)</span></span>  | <span data-ttu-id="4cb81-121">Office 网页版\*</span><span class="sxs-lookup"><span data-stu-id="4cb81-121">Office on the web\*</span></span>  |  <span data-ttu-id="4cb81-122">Office Online Server</span><span class="sxs-lookup"><span data-stu-id="4cb81-122">Office Online Server</span></span>  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="4cb81-123">RibbonApi 1.1</span><span class="sxs-lookup"><span data-stu-id="4cb81-123">RibbonApi 1.1</span></span>  | <span data-ttu-id="4cb81-124">不适用</span><span class="sxs-lookup"><span data-stu-id="4cb81-124">N/A</span></span> | <span data-ttu-id="4cb81-125">不适用</span><span class="sxs-lookup"><span data-stu-id="4cb81-125">N/A</span></span> | <span data-ttu-id="4cb81-126">请参阅支持</span><span class="sxs-lookup"><span data-stu-id="4cb81-126">See support</span></span><br><span data-ttu-id="4cb81-127">部分如下</span><span class="sxs-lookup"><span data-stu-id="4cb81-127">section below</span></span> | <span data-ttu-id="4cb81-128">无</span><span class="sxs-lookup"><span data-stu-id="4cb81-128">N/A</span></span> | <span data-ttu-id="4cb81-129">16.38</span><span class="sxs-lookup"><span data-stu-id="4cb81-129">16.38</span></span> | <span data-ttu-id="4cb81-130">2020 年 11 月</span><span class="sxs-lookup"><span data-stu-id="4cb81-130">November, 2020</span></span> | <span data-ttu-id="4cb81-131">无</span><span class="sxs-lookup"><span data-stu-id="4cb81-131">N/A</span></span>|

> <span data-ttu-id="4cb81-132">**&#42;** 功能区 API 仅在 Excel 上受支持，并且需要 Microsoft 365 订阅。</span><span class="sxs-lookup"><span data-stu-id="4cb81-132">**&#42;** The Ribbon API is supported only on Excel and it requires Microsoft 365 subscription.</span></span>

## <a name="office-on-windows-subscription-support"></a><span data-ttu-id="4cb81-133">Windows 版 Office (订阅) 支持</span><span class="sxs-lookup"><span data-stu-id="4cb81-133">Office on Windows (subscription) support</span></span>

<span data-ttu-id="4cb81-134">要求集在消费者频道版本 2006 (版本 13001.20498 或) 。</span><span class="sxs-lookup"><span data-stu-id="4cb81-134">The requirement set is supported in the Consumer Channel version 2006 (build, 13001.20498 or greater).</span></span> <span data-ttu-id="4cb81-135">对于 Windows 版 Office，2020 Semi-Annual 2020 年 7 月 14 日版和每月企业频道版本也支持此功能。</span><span class="sxs-lookup"><span data-stu-id="4cb81-135">For Office on Windows the feature is also supported in the Semi-Annual Channel and Monthly Enterprise Channel builds available July 14th, 2020 or later.</span></span> <span data-ttu-id="4cb81-136">每个频道支持的最低版本如下：</span><span class="sxs-lookup"><span data-stu-id="4cb81-136">The minimum supported builds for each channel are as follows:</span></span>  

|<span data-ttu-id="4cb81-137">频道</span><span class="sxs-lookup"><span data-stu-id="4cb81-137">Channel</span></span> | <span data-ttu-id="4cb81-138">版本</span><span class="sxs-lookup"><span data-stu-id="4cb81-138">Version</span></span> | <span data-ttu-id="4cb81-139">内部版本</span><span class="sxs-lookup"><span data-stu-id="4cb81-139">Build</span></span>|
|:-----|:-----|:-----|
|<span data-ttu-id="4cb81-140">当前频道</span><span class="sxs-lookup"><span data-stu-id="4cb81-140">Current Channel</span></span> | <span data-ttu-id="4cb81-141">2006 或更大</span><span class="sxs-lookup"><span data-stu-id="4cb81-141">2006 or greater</span></span> | <span data-ttu-id="4cb81-142">20266.20266 或更大</span><span class="sxs-lookup"><span data-stu-id="4cb81-142">20266.20266 or greater</span></span>|
|<span data-ttu-id="4cb81-143">每月企业频道</span><span class="sxs-lookup"><span data-stu-id="4cb81-143">Monthly Enterprise Channel</span></span> | <span data-ttu-id="4cb81-144">2005 或更大</span><span class="sxs-lookup"><span data-stu-id="4cb81-144">2005 or greater</span></span> | <span data-ttu-id="4cb81-145">12827.20538 或更大</span><span class="sxs-lookup"><span data-stu-id="4cb81-145">12827.20538 or greater</span></span>|
|<span data-ttu-id="4cb81-146">每月企业频道</span><span class="sxs-lookup"><span data-stu-id="4cb81-146">Monthly Enterprise Channel</span></span> | <span data-ttu-id="4cb81-147">2004</span><span class="sxs-lookup"><span data-stu-id="4cb81-147">2004</span></span> | <span data-ttu-id="4cb81-148">12730.20602 或更大</span><span class="sxs-lookup"><span data-stu-id="4cb81-148">12730.20602 or greater</span></span>|
|<span data-ttu-id="4cb81-149">半年企业频道</span><span class="sxs-lookup"><span data-stu-id="4cb81-149">Semi-Annual Enterprise Channel</span></span> | <span data-ttu-id="4cb81-150">2002 或更大</span><span class="sxs-lookup"><span data-stu-id="4cb81-150">2002 or greater</span></span> | <span data-ttu-id="4cb81-151">12527.20880 或更大</span><span class="sxs-lookup"><span data-stu-id="4cb81-151">12527.20880 or greater</span></span>|

## <a name="more-information"></a><span data-ttu-id="4cb81-152">更多信息</span><span class="sxs-lookup"><span data-stu-id="4cb81-152">More information</span></span>

<span data-ttu-id="4cb81-153">若要详细了解版本、内部版本号和 Office Online Server，请参阅：</span><span class="sxs-lookup"><span data-stu-id="4cb81-153">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

- [<span data-ttu-id="4cb81-154">Microsoft 365 客户端更新频道版本的版本号和内部版本号</span><span class="sxs-lookup"><span data-stu-id="4cb81-154">Version and build numbers of update channel releases for Microsoft 365 clients</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="4cb81-155">使用的是哪一版 Office？</span><span class="sxs-lookup"><span data-stu-id="4cb81-155">What version of Office am I using?</span></span>](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [<span data-ttu-id="4cb81-156">在哪里可以找到 Microsoft 365 客户端应用程序的版本号和内部版本号</span><span class="sxs-lookup"><span data-stu-id="4cb81-156">Where you can find the version and build number for a Microsoft 365 client application</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="4cb81-157">Office Online Server 概述</span><span class="sxs-lookup"><span data-stu-id="4cb81-157">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

> [!NOTE]
> <span data-ttu-id="4cb81-158">**由于清单中尚不支持 RibbonApi 1.1** 要求集，因此无法在清单的部分中指定 `<Requirements>` 它。</span><span class="sxs-lookup"><span data-stu-id="4cb81-158">The **RibbonApi 1.1** requirement set is not yet supported in the manifest, so you cannot specify it in the manifest's `<Requirements>` section.</span></span>


## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="4cb81-159">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="4cb81-159">Office Common API requirement sets</span></span>

<span data-ttu-id="4cb81-160">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="4cb81-160">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="ribbon-api-11"></a><span data-ttu-id="4cb81-161">功能区 API 1.1</span><span class="sxs-lookup"><span data-stu-id="4cb81-161">Ribbon API 1.1</span></span>

<span data-ttu-id="4cb81-162">功能区 API 1.1 是 API 的第一个版本。</span><span class="sxs-lookup"><span data-stu-id="4cb81-162">The Ribbon API 1.1 is the first version of the API.</span></span> <span data-ttu-id="4cb81-163">有关 API 的详细信息，请参阅 [Office.ribbon ](/javascript/api/office/office.ribbon) 参考主题。</span><span class="sxs-lookup"><span data-stu-id="4cb81-163">For details about the API, see the [Office.ribbon ](/javascript/api/office/office.ribbon) reference topic.</span></span>

## <a name="see-also"></a><span data-ttu-id="4cb81-164">另请参阅</span><span class="sxs-lookup"><span data-stu-id="4cb81-164">See also</span></span>

- [<span data-ttu-id="4cb81-165">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="4cb81-165">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="4cb81-166">指定 Office 应用程序和 API 要求</span><span class="sxs-lookup"><span data-stu-id="4cb81-166">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="4cb81-167">Office 加载项 XML 清单</span><span class="sxs-lookup"><span data-stu-id="4cb81-167">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)