---
title: Dialog API 要求集
description: ''
ms.date: 06/20/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 3135c65120248194603b91510450519f106e0ad1
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127119"
---
# <a name="dialog-api-requirement-sets"></a><span data-ttu-id="3987e-102">Dialog API 要求集</span><span class="sxs-lookup"><span data-stu-id="3987e-102">Dialog API requirement sets</span></span>

<span data-ttu-id="3987e-p101">要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="3987e-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="3987e-p102">Office 外接程序在多个 Office 版本中运行。下表列出了 Dialog API 要求集、支持该要求集的 Office 主机应用程序，以及 Office 应用程序的内部版本或版本号。</span><span class="sxs-lookup"><span data-stu-id="3987e-p102">Office Add-ins run across multiple versions of Office. The following table lists the Dialog API requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="3987e-108">要求集</span><span class="sxs-lookup"><span data-stu-id="3987e-108">Requirement set</span></span>  | <span data-ttu-id="3987e-109">Windows 版 Office 2013</span><span class="sxs-lookup"><span data-stu-id="3987e-109">Office 2013 on Windows</span></span><br><span data-ttu-id="3987e-110">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3987e-110">(one-time purchase)</span></span> | <span data-ttu-id="3987e-111">Windows 上的 Office 2016 或更高版本</span><span class="sxs-lookup"><span data-stu-id="3987e-111">Office 2016 or later on Windows</span></span><br><span data-ttu-id="3987e-112">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3987e-112">(one-time purchase)</span></span>   | <span data-ttu-id="3987e-113">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="3987e-113">Office on Windows</span></span><br><span data-ttu-id="3987e-114">(连接到 Office 365 订阅)</span><span class="sxs-lookup"><span data-stu-id="3987e-114">(connected to Office 365 subscription)</span></span> |  <span data-ttu-id="3987e-115">IPad 上的 Office</span><span class="sxs-lookup"><span data-stu-id="3987e-115">Office on iPad</span></span><br><span data-ttu-id="3987e-116">(连接到 Office 365 订阅)</span><span class="sxs-lookup"><span data-stu-id="3987e-116">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="3987e-117">Mac 上的 Office</span><span class="sxs-lookup"><span data-stu-id="3987e-117">Office on Mac</span></span><br><span data-ttu-id="3987e-118">(连接到 Office 365 订阅)</span><span class="sxs-lookup"><span data-stu-id="3987e-118">(connected to Office 365 subscription)</span></span>  | <span data-ttu-id="3987e-119">网上的 Office</span><span class="sxs-lookup"><span data-stu-id="3987e-119">Office on the web</span></span>  |  <span data-ttu-id="3987e-120">Office Online Server</span><span class="sxs-lookup"><span data-stu-id="3987e-120">Office Online Server</span></span>  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="3987e-121">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="3987e-121">DialogApi 1.1</span></span>  | <span data-ttu-id="3987e-122">生成号 15.0.4855.1000 或更高版本</span><span class="sxs-lookup"><span data-stu-id="3987e-122">Build 15.0.4855.1000 or later</span></span> | <span data-ttu-id="3987e-123">生成号 16.0.4390.1000 或更高版本</span><span class="sxs-lookup"><span data-stu-id="3987e-123">Build 16.0.4390.1000 or later</span></span> | <span data-ttu-id="3987e-124">版本 1602（生成号 6741.0000）或更高版本</span><span class="sxs-lookup"><span data-stu-id="3987e-124">Version 1602 (Build 6741.0000) or later</span></span> | <span data-ttu-id="3987e-125">1.22 或更高版本</span><span class="sxs-lookup"><span data-stu-id="3987e-125">1.22 or later</span></span> | <span data-ttu-id="3987e-126">15.20 或更高版本</span><span class="sxs-lookup"><span data-stu-id="3987e-126">15.20 or later</span></span>| <span data-ttu-id="3987e-127">2017 年 1 月</span><span class="sxs-lookup"><span data-stu-id="3987e-127">January 2017</span></span> | <span data-ttu-id="3987e-128">版本 1608（生成号 7601.6800）或更高版本</span><span class="sxs-lookup"><span data-stu-id="3987e-128">Version 1608 (Build 7601.6800) or later</span></span>|

<span data-ttu-id="3987e-129">若要详细了解版本、内部版本号和 Office Online Server，请参阅：</span><span class="sxs-lookup"><span data-stu-id="3987e-129">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

- [<span data-ttu-id="3987e-130">更新频道发布的 Office 365 客户端版本号和内部版本号</span><span class="sxs-lookup"><span data-stu-id="3987e-130">Version and build numbers of update channel releases for Office 365 clients</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="3987e-131">使用的是哪一版 Office？</span><span class="sxs-lookup"><span data-stu-id="3987e-131">What version of Office am I using?</span></span>](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [<span data-ttu-id="3987e-132">在哪里可以找到 Office 365 客户端应用程序的版本号和内部版本号</span><span class="sxs-lookup"><span data-stu-id="3987e-132">Where you can find the version and build number for an Office 365 client application</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="3987e-133">Office Online Server 概述</span><span class="sxs-lookup"><span data-stu-id="3987e-133">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="3987e-134">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="3987e-134">Office Common API requirement sets</span></span>

<span data-ttu-id="3987e-135">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="3987e-135">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="dialog-api-11"></a><span data-ttu-id="3987e-136">Dialog API 1.1</span><span class="sxs-lookup"><span data-stu-id="3987e-136">Dialog API 1.1</span></span>

<span data-ttu-id="3987e-137">Dialog API 1.1 是首版 API。</span><span class="sxs-lookup"><span data-stu-id="3987e-137">The Dialog API 1.1 is the first version of the API.</span></span> <span data-ttu-id="3987e-138">有关 API 的详细信息，请参阅 [Dialog API](/javascript/api/office/office.ui) 参考主题。</span><span class="sxs-lookup"><span data-stu-id="3987e-138">For details about the API, see the [Dialog API ](/javascript/api/office/office.ui) reference topic.</span></span>

## <a name="see-also"></a><span data-ttu-id="3987e-139">另请参阅</span><span class="sxs-lookup"><span data-stu-id="3987e-139">See also</span></span>

- [<span data-ttu-id="3987e-140">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="3987e-140">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="3987e-141">指定 Office 主机和 API 要求</span><span class="sxs-lookup"><span data-stu-id="3987e-141">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="3987e-142">Office 外接程序 XML 清单</span><span class="sxs-lookup"><span data-stu-id="3987e-142">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
