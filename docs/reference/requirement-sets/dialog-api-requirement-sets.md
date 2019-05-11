---
title: Dialog API 要求集
description: ''
ms.date: 05/08/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: f6f0b0184736bfd0f6b417198ade4c621d8d8b6b
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952199"
---
# <a name="dialog-api-requirement-sets"></a><span data-ttu-id="a55af-102">Dialog API 要求集</span><span class="sxs-lookup"><span data-stu-id="a55af-102">Dialog API requirement sets</span></span>

<span data-ttu-id="a55af-p101">要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="a55af-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="a55af-p102">Office 外接程序在多个 Office 版本中运行。下表列出了 Dialog API 要求集、支持该要求集的 Office 主机应用程序，以及 Office 应用程序的内部版本或版本号。</span><span class="sxs-lookup"><span data-stu-id="a55af-p102">Office Add-ins run across multiple versions of Office. The following table lists the Dialog API requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="a55af-108">要求集</span><span class="sxs-lookup"><span data-stu-id="a55af-108">Requirement set</span></span>  | <span data-ttu-id="a55af-109">Windows 上的 Office 2013</span><span class="sxs-lookup"><span data-stu-id="a55af-109">Office 2013 on Windows</span></span><br><span data-ttu-id="a55af-110">(一次性购买)</span><span class="sxs-lookup"><span data-stu-id="a55af-110">(one-time purchase)</span></span> | <span data-ttu-id="a55af-111">Windows 上的 Office 2016 或更高版本</span><span class="sxs-lookup"><span data-stu-id="a55af-111">Office 2016 or later on Windows</span></span><br><span data-ttu-id="a55af-112">(一次性购买)</span><span class="sxs-lookup"><span data-stu-id="a55af-112">(one-time purchase)</span></span>   | <span data-ttu-id="a55af-113">Windows 上的 Office</span><span class="sxs-lookup"><span data-stu-id="a55af-113">Office on Windows</span></span><br><span data-ttu-id="a55af-114">(已连接到 Office 365)</span><span class="sxs-lookup"><span data-stu-id="a55af-114">(connected to Office 365)</span></span> |  <span data-ttu-id="a55af-115">Office for iPad</span><span class="sxs-lookup"><span data-stu-id="a55af-115">Office for iPad</span></span><br><span data-ttu-id="a55af-116">(已连接到 Office 365)</span><span class="sxs-lookup"><span data-stu-id="a55af-116">(connected to Office 365)</span></span>  |  <span data-ttu-id="a55af-117">Office for Mac</span><span class="sxs-lookup"><span data-stu-id="a55af-117">Office for Mac</span></span><br><span data-ttu-id="a55af-118">(已连接到 Office 365)</span><span class="sxs-lookup"><span data-stu-id="a55af-118">(connected to Office 365)</span></span>  | <span data-ttu-id="a55af-119">Office Online</span><span class="sxs-lookup"><span data-stu-id="a55af-119">Office Online</span></span>  |  <span data-ttu-id="a55af-120">Office Online Server</span><span class="sxs-lookup"><span data-stu-id="a55af-120">Office Online Server</span></span>  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="a55af-121">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="a55af-121">DialogApi 1.1</span></span>  | <span data-ttu-id="a55af-122">生成号 15.0.4855.1000 或更高版本</span><span class="sxs-lookup"><span data-stu-id="a55af-122">Build 15.0.4855.1000 or later</span></span> | <span data-ttu-id="a55af-123">生成号 16.0.4390.1000 或更高版本</span><span class="sxs-lookup"><span data-stu-id="a55af-123">Build 16.0.4390.1000 or later</span></span> | <span data-ttu-id="a55af-124">版本 1602（生成号 6741.0000）或更高版本</span><span class="sxs-lookup"><span data-stu-id="a55af-124">Version 1602 (Build 6741.0000) or later</span></span> | <span data-ttu-id="a55af-125">1.22 或更高版本</span><span class="sxs-lookup"><span data-stu-id="a55af-125">1.22 or later</span></span> | <span data-ttu-id="a55af-126">15.20 或更高版本</span><span class="sxs-lookup"><span data-stu-id="a55af-126">15.20 or later</span></span>| <span data-ttu-id="a55af-127">2017 年 1 月</span><span class="sxs-lookup"><span data-stu-id="a55af-127">January 2017</span></span> | <span data-ttu-id="a55af-128">版本 1608（生成号 7601.6800）或更高版本</span><span class="sxs-lookup"><span data-stu-id="a55af-128">Version 1608 (Build 7601.6800) or later</span></span>|

<span data-ttu-id="a55af-129">若要详细了解版本、内部版本号和 Office Online Server，请参阅：</span><span class="sxs-lookup"><span data-stu-id="a55af-129">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

- [<span data-ttu-id="a55af-130">更新频道发布的 Office 365 客户端版本号和内部版本号</span><span class="sxs-lookup"><span data-stu-id="a55af-130">Version and build numbers of update channel releases for Office 365 clients</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="a55af-131">使用的是哪一版 Office？</span><span class="sxs-lookup"><span data-stu-id="a55af-131">What version of Office am I using?</span></span>](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [<span data-ttu-id="a55af-132">在哪里可以找到 Office 365 客户端应用程序的版本号和内部版本号</span><span class="sxs-lookup"><span data-stu-id="a55af-132">Where you can find the version and build number for an Office 365 client application</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="a55af-133">Office Online Server 概述</span><span class="sxs-lookup"><span data-stu-id="a55af-133">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="a55af-134">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="a55af-134">Office Common API requirement sets</span></span>

<span data-ttu-id="a55af-135">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="a55af-135">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="dialog-api-11"></a><span data-ttu-id="a55af-136">Dialog API 1.1</span><span class="sxs-lookup"><span data-stu-id="a55af-136">Dialog API 1.1</span></span>

<span data-ttu-id="a55af-137">Dialog API 1.1 是首版 API。</span><span class="sxs-lookup"><span data-stu-id="a55af-137">The Dialog API 1.1 is the first version of the API.</span></span> <span data-ttu-id="a55af-138">有关 API 的详细信息，请参阅 [Dialog API](/javascript/api/office/office.ui) 参考主题。</span><span class="sxs-lookup"><span data-stu-id="a55af-138">For details about the API, see the [Dialog API ](/javascript/api/office/office.ui) reference topic.</span></span>

## <a name="see-also"></a><span data-ttu-id="a55af-139">另请参阅</span><span class="sxs-lookup"><span data-stu-id="a55af-139">See also</span></span>

- [<span data-ttu-id="a55af-140">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="a55af-140">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="a55af-141">指定 Office 主机和 API 要求</span><span class="sxs-lookup"><span data-stu-id="a55af-141">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="a55af-142">Office 外接程序 XML 清单</span><span class="sxs-lookup"><span data-stu-id="a55af-142">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
