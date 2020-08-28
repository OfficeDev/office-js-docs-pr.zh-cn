---
title: 功能区 API 要求集
description: 指定哪些 Office 平台和生成支持动态功能区 Api。
ms.date: 08/26/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: f734931817111ce52f779946e1f983ecc9238d3a
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293490"
---
# <a name="ribbon-api-requirement-sets"></a><span data-ttu-id="88e4e-103">功能区 API 要求集</span><span class="sxs-lookup"><span data-stu-id="88e4e-103">Ribbon API requirement sets</span></span>

<span data-ttu-id="88e4e-104">要求集是指各组已命名的 API 成员。</span><span class="sxs-lookup"><span data-stu-id="88e4e-104">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="88e4e-105">Office 外接程序使用清单中指定的要求集或使用运行时检查来确定 Office 应用程序是否支持加载项所需的 Api。</span><span class="sxs-lookup"><span data-stu-id="88e4e-105">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs.</span></span> <span data-ttu-id="88e4e-106">有关详细信息，请参阅 [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="88e4e-106">For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="88e4e-107">功能区 API 集支持编程控制何时自定义外接程序命令 (即启用和禁用自定义功能区按钮和菜单项) 。</span><span class="sxs-lookup"><span data-stu-id="88e4e-107">The Ribbon API set supports programmatic control of when custom Add-in Commands (that is, custom ribbon buttons and menu items) are enabled and disabled.</span></span>

<span data-ttu-id="88e4e-108">Office 外接程序在多个 Office 版本中运行。</span><span class="sxs-lookup"><span data-stu-id="88e4e-108">Office Add-ins run across multiple versions of Office.</span></span> <span data-ttu-id="88e4e-109">下表列出了功能区 API 要求集、支持该要求集的 Office 客户端应用程序，以及 Office 应用程序的内部版本号或版本号。</span><span class="sxs-lookup"><span data-stu-id="88e4e-109">The following table lists the Ribbon API requirement sets, the Office client applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="88e4e-110">要求集</span><span class="sxs-lookup"><span data-stu-id="88e4e-110">Requirement set</span></span>  | <span data-ttu-id="88e4e-111">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="88e4e-111">Office 2013 on Windows</span></span><br><span data-ttu-id="88e4e-112">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="88e4e-112">(one-time purchase)</span></span> | <span data-ttu-id="88e4e-113">Windows 上的 Office 2016 或更高版本</span><span class="sxs-lookup"><span data-stu-id="88e4e-113">Office 2016 or later on Windows</span></span><br><span data-ttu-id="88e4e-114">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="88e4e-114">(one-time purchase)</span></span>   | <span data-ttu-id="88e4e-115">Windows 版 Office\*</span><span class="sxs-lookup"><span data-stu-id="88e4e-115">Office on Windows\*</span></span><br><span data-ttu-id="88e4e-116">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="88e4e-116">(connected to a Microsoft 365 subscription)</span></span> |  <span data-ttu-id="88e4e-117">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="88e4e-117">Office on iPad</span></span><br><span data-ttu-id="88e4e-118">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="88e4e-118">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="88e4e-119">Mac 版 Office\*</span><span class="sxs-lookup"><span data-stu-id="88e4e-119">Office on Mac\*</span></span><br><span data-ttu-id="88e4e-120">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="88e4e-120">(connected to a Microsoft 365 subscription)</span></span>  | <span data-ttu-id="88e4e-121">Office 网页版\*</span><span class="sxs-lookup"><span data-stu-id="88e4e-121">Office on the web\*</span></span>  |  <span data-ttu-id="88e4e-122">Office Online Server</span><span class="sxs-lookup"><span data-stu-id="88e4e-122">Office Online Server</span></span>  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="88e4e-123">RibbonApi 1。1</span><span class="sxs-lookup"><span data-stu-id="88e4e-123">RibbonApi 1.1</span></span>  | <span data-ttu-id="88e4e-124">不适用</span><span class="sxs-lookup"><span data-stu-id="88e4e-124">N/A</span></span> | <span data-ttu-id="88e4e-125">不适用</span><span class="sxs-lookup"><span data-stu-id="88e4e-125">N/A</span></span> | <span data-ttu-id="88e4e-126">请参阅支持</span><span class="sxs-lookup"><span data-stu-id="88e4e-126">See support</span></span><br><span data-ttu-id="88e4e-127">部分</span><span class="sxs-lookup"><span data-stu-id="88e4e-127">section below</span></span> | <span data-ttu-id="88e4e-128">无</span><span class="sxs-lookup"><span data-stu-id="88e4e-128">N/A</span></span> | <span data-ttu-id="88e4e-129">16.38</span><span class="sxs-lookup"><span data-stu-id="88e4e-129">16.38</span></span> | <span data-ttu-id="88e4e-130">即将推出</span><span class="sxs-lookup"><span data-stu-id="88e4e-130">Coming soon</span></span> | <span data-ttu-id="88e4e-131">无</span><span class="sxs-lookup"><span data-stu-id="88e4e-131">N/A</span></span>|

> <span data-ttu-id="88e4e-132">**&#42;** 仅在 Excel 中支持功能区 API，并且它需要 Microsoft 365 订阅。</span><span class="sxs-lookup"><span data-stu-id="88e4e-132">**&#42;** The Ribbon API is supported only on Excel and it requires Microsoft 365 subscription.</span></span> 

## <a name="office-on-windows-subscription-support"></a><span data-ttu-id="88e4e-133">Office on Windows (订阅) 支持</span><span class="sxs-lookup"><span data-stu-id="88e4e-133">Office on Windows (subscription) support</span></span>

<span data-ttu-id="88e4e-134">使用者通道版本 2006 (版本13001.20498 或更高版本中支持的要求集) 。</span><span class="sxs-lookup"><span data-stu-id="88e4e-134">The requirement set is supported in the Consumer Channel version 2006 (build, 13001.20498 or greater).</span></span> <span data-ttu-id="88e4e-135">对于 Windows 上的 Office，在半年频道和每月14月14日（2020或更高版本）中也支持此功能。</span><span class="sxs-lookup"><span data-stu-id="88e4e-135">For Office on Windows the feature is also supported in the Semi-Annual Channel and Monthly Enterprise Channel builds available July 14th, 2020 or later.</span></span> <span data-ttu-id="88e4e-136">每个频道支持的最低版本如下所示：</span><span class="sxs-lookup"><span data-stu-id="88e4e-136">The minimum supported builds for each channel are as follows:</span></span>  

|<span data-ttu-id="88e4e-137">频道</span><span class="sxs-lookup"><span data-stu-id="88e4e-137">Channel</span></span> | <span data-ttu-id="88e4e-138">版本</span><span class="sxs-lookup"><span data-stu-id="88e4e-138">Version</span></span> | <span data-ttu-id="88e4e-139">内部版本</span><span class="sxs-lookup"><span data-stu-id="88e4e-139">Build</span></span>|
|:-----|:-----|:-----|
|<span data-ttu-id="88e4e-140">当前频道</span><span class="sxs-lookup"><span data-stu-id="88e4e-140">Current Channel</span></span> | <span data-ttu-id="88e4e-141">2006或更高版本</span><span class="sxs-lookup"><span data-stu-id="88e4e-141">2006 or greater</span></span> | <span data-ttu-id="88e4e-142">20266.20266 或更高版本</span><span class="sxs-lookup"><span data-stu-id="88e4e-142">20266.20266 or greater</span></span>|
|<span data-ttu-id="88e4e-143">月度企业版频道</span><span class="sxs-lookup"><span data-stu-id="88e4e-143">Monthly Enterprise Channel</span></span> | <span data-ttu-id="88e4e-144">2005或更高版本</span><span class="sxs-lookup"><span data-stu-id="88e4e-144">2005 or greater</span></span> | <span data-ttu-id="88e4e-145">12827.20538 或更高版本</span><span class="sxs-lookup"><span data-stu-id="88e4e-145">12827.20538 or greater</span></span>|
|<span data-ttu-id="88e4e-146">每月企业频道</span><span class="sxs-lookup"><span data-stu-id="88e4e-146">Monthly Enterprise Channel</span></span> | <span data-ttu-id="88e4e-147">2004</span><span class="sxs-lookup"><span data-stu-id="88e4e-147">2004</span></span> | <span data-ttu-id="88e4e-148">12730.20602 或更高版本</span><span class="sxs-lookup"><span data-stu-id="88e4e-148">12730.20602 or greater</span></span>|
|<span data-ttu-id="88e4e-149">半年企业频道</span><span class="sxs-lookup"><span data-stu-id="88e4e-149">Semi-Annual Enterprise Channel</span></span> | <span data-ttu-id="88e4e-150">2002或更高版本</span><span class="sxs-lookup"><span data-stu-id="88e4e-150">2002 or greater</span></span> | <span data-ttu-id="88e4e-151">12527.20880 或更高版本</span><span class="sxs-lookup"><span data-stu-id="88e4e-151">12527.20880 or greater</span></span>|

## <a name="more-information"></a><span data-ttu-id="88e4e-152">更多信息</span><span class="sxs-lookup"><span data-stu-id="88e4e-152">More information</span></span>

<span data-ttu-id="88e4e-153">若要详细了解版本、内部版本号和 Office Online Server，请参阅：</span><span class="sxs-lookup"><span data-stu-id="88e4e-153">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

- [<span data-ttu-id="88e4e-154">适用于 Microsoft 365 客户端的更新通道版本和内部版本号</span><span class="sxs-lookup"><span data-stu-id="88e4e-154">Version and build numbers of update channel releases for Microsoft 365 clients</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="88e4e-155">使用的是哪一版 Office？</span><span class="sxs-lookup"><span data-stu-id="88e4e-155">What version of Office am I using?</span></span>](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [<span data-ttu-id="88e4e-156">在哪里可以找到 Microsoft 365 客户端应用程序的版本和内部版本号</span><span class="sxs-lookup"><span data-stu-id="88e4e-156">Where you can find the version and build number for a Microsoft 365 client application</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="88e4e-157">Office Online Server 概述</span><span class="sxs-lookup"><span data-stu-id="88e4e-157">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

> [!NOTE]
> <span data-ttu-id="88e4e-158">**RibbonApi 1.1**要求集在清单中尚不受支持，因此不能在清单的部分中指定它 `<Requirements>` 。</span><span class="sxs-lookup"><span data-stu-id="88e4e-158">The **RibbonApi 1.1** requirement set is not yet supported in the manifest, so you cannot specify it in the manifest's `<Requirements>` section.</span></span>


## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="88e4e-159">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="88e4e-159">Office Common API requirement sets</span></span>

<span data-ttu-id="88e4e-160">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="88e4e-160">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="ribbon-api-11"></a><span data-ttu-id="88e4e-161">功能区 API 1。1</span><span class="sxs-lookup"><span data-stu-id="88e4e-161">Ribbon API 1.1</span></span>

<span data-ttu-id="88e4e-162">功能区 API 1.1 是 API 的第一个版本。</span><span class="sxs-lookup"><span data-stu-id="88e4e-162">The Ribbon API 1.1 is the first version of the API.</span></span> <span data-ttu-id="88e4e-163">有关 API 的详细信息，请参阅 " [Office. 功能区 ](/javascript/api/office/office.ribbon) 参考" 主题。</span><span class="sxs-lookup"><span data-stu-id="88e4e-163">For details about the API, see the [Office.ribbon ](/javascript/api/office/office.ribbon) reference topic.</span></span>

## <a name="see-also"></a><span data-ttu-id="88e4e-164">另请参阅</span><span class="sxs-lookup"><span data-stu-id="88e4e-164">See also</span></span>

- [<span data-ttu-id="88e4e-165">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="88e4e-165">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="88e4e-166">指定 Office 应用程序和 API 要求</span><span class="sxs-lookup"><span data-stu-id="88e4e-166">Specify Office applications and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="88e4e-167">Office 外接程序 XML 清单</span><span class="sxs-lookup"><span data-stu-id="88e4e-167">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
