---
title: Dialog API 要求集
description: ''
ms.date: 03/19/2019
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: ebbd10e65894a7d038e54ffbaac20c973adf4a9f
ms.sourcegitcommit: c5daedf017c6dd5ab0c13607589208c3f3627354
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/20/2019
ms.locfileid: "30691130"
---
# <a name="dialog-api-requirement-sets"></a><span data-ttu-id="aa829-102">Dialog API 要求集</span><span class="sxs-lookup"><span data-stu-id="aa829-102">Dialog API requirement sets</span></span>

<span data-ttu-id="aa829-p101">要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="aa829-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="aa829-p102">Office 外接程序在多个 Office 版本中运行。下表列出了 Dialog API 要求集、支持该要求集的 Office 主机应用程序，以及 Office 应用程序的内部版本或版本号。</span><span class="sxs-lookup"><span data-stu-id="aa829-p102">Office Add-ins run across multiple versions of Office. The following table lists the Dialog API requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="aa829-108">要求集</span><span class="sxs-lookup"><span data-stu-id="aa829-108">Requirement set</span></span>  | <span data-ttu-id="aa829-109">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="aa829-109">Office 2013 for Windows</span></span> | <span data-ttu-id="aa829-110">适用于 Windows 的 Office 2016 或更高版本</span><span class="sxs-lookup"><span data-stu-id="aa829-110">Office 2016 or later for Windows</span></span>   | <span data-ttu-id="aa829-111">Office 365 for Windows</span><span class="sxs-lookup"><span data-stu-id="aa829-111">Office 365 for Windows</span></span> |  <span data-ttu-id="aa829-112">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="aa829-112">Office 365 for iPad</span></span>  |  <span data-ttu-id="aa829-113">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="aa829-113">Office 365 for Mac</span></span>  | <span data-ttu-id="aa829-114">Office Online</span><span class="sxs-lookup"><span data-stu-id="aa829-114">Office Online</span></span>  |  <span data-ttu-id="aa829-115">Office Online Server</span><span class="sxs-lookup"><span data-stu-id="aa829-115">Office Online Server</span></span>  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="aa829-116">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="aa829-116">DialogApi 1.1</span></span>  | <span data-ttu-id="aa829-117">生成号 15.0.4855.1000 或更高版本</span><span class="sxs-lookup"><span data-stu-id="aa829-117">Build 15.0.4855.1000 or later</span></span> | <span data-ttu-id="aa829-118">生成号 16.0.4390.1000 或更高版本</span><span class="sxs-lookup"><span data-stu-id="aa829-118">Build 16.0.4390.1000 or later</span></span> | <span data-ttu-id="aa829-119">版本 1602（生成号 6741.0000）或更高版本</span><span class="sxs-lookup"><span data-stu-id="aa829-119">Version 1602 (Build 6741.0000) or later</span></span> | <span data-ttu-id="aa829-120">1.22 或更高版本</span><span class="sxs-lookup"><span data-stu-id="aa829-120">1.22 or later</span></span> | <span data-ttu-id="aa829-121">15.20 或更高版本</span><span class="sxs-lookup"><span data-stu-id="aa829-121">15.20 or later</span></span>| <span data-ttu-id="aa829-122">2017 年 1 月</span><span class="sxs-lookup"><span data-stu-id="aa829-122">January 2017</span></span> | <span data-ttu-id="aa829-123">版本 1608（生成号 7601.6800）或更高版本</span><span class="sxs-lookup"><span data-stu-id="aa829-123">Version 1608 (Build 7601.6800) or later</span></span>|

<span data-ttu-id="aa829-124">若要详细了解版本、生成号和 Office Online Server，请参阅：</span><span class="sxs-lookup"><span data-stu-id="aa829-124">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

- [<span data-ttu-id="aa829-125">更新频道发布的 Office 365 客户端版本号和内部版本号</span><span class="sxs-lookup"><span data-stu-id="aa829-125">Version and build numbers of update channel releases for Office 365 clients</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="aa829-126">使用的是哪一版 Office？</span><span class="sxs-lookup"><span data-stu-id="aa829-126">What version of Office am I using?</span></span>](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [<span data-ttu-id="aa829-127">在哪里可以找到 Office 365 客户端应用程序的版本号和内部版本号</span><span class="sxs-lookup"><span data-stu-id="aa829-127">Where you can find the version and build number for an Office 365 client application</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="aa829-128">Office Online Server 概述</span><span class="sxs-lookup"><span data-stu-id="aa829-128">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="aa829-129">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="aa829-129">Office Common API requirement sets</span></span>

<span data-ttu-id="aa829-130">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="aa829-130">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="dialog-api-11"></a><span data-ttu-id="aa829-131">Dialog API 1.1</span><span class="sxs-lookup"><span data-stu-id="aa829-131">Dialog API 1.1</span></span>

<span data-ttu-id="aa829-132">Dialog API 1.1 是首版 API。</span><span class="sxs-lookup"><span data-stu-id="aa829-132">The Dialog API 1.1 is the first version of the API.</span></span> <span data-ttu-id="aa829-133">有关 API 的详细信息，请参阅 [Dialog API](/javascript/api/office/office.ui) 参考主题。</span><span class="sxs-lookup"><span data-stu-id="aa829-133">For details about the API, see the [Dialog API ](/javascript/api/office/office.ui) reference topic.</span></span>

## <a name="see-also"></a><span data-ttu-id="aa829-134">另请参阅</span><span class="sxs-lookup"><span data-stu-id="aa829-134">See also</span></span>

- [<span data-ttu-id="aa829-135">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="aa829-135">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="aa829-136">指定 Office 主机和 API 要求</span><span class="sxs-lookup"><span data-stu-id="aa829-136">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="aa829-137">Office 外接程序 XML 清单</span><span class="sxs-lookup"><span data-stu-id="aa829-137">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
