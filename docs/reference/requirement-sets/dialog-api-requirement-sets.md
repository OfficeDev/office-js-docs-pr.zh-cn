---
title: Dialog API 要求集
description: ''
ms.date: 10/09/2018
ms.openlocfilehash: ad0d472ebdcbdb9d61e78f6bdc9bfe7c08311cd7
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432653"
---
# <a name="dialog-api-requirement-sets"></a><span data-ttu-id="7d0f0-102">Dialog API 要求集</span><span class="sxs-lookup"><span data-stu-id="7d0f0-102">Dialog API requirement sets</span></span>

<span data-ttu-id="7d0f0-p101">要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="7d0f0-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="7d0f0-p102">Office 外接程序在多个 Office 版本中运行。下表列出了 Dialog API 要求集、支持该要求集的 Office 主机应用程序，以及 Office 应用程序的内部版本或版本号。</span><span class="sxs-lookup"><span data-stu-id="7d0f0-p102">Office Add-ins run across multiple versions of Office. The following table lists the Dialog API requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="7d0f0-108">要求集</span><span class="sxs-lookup"><span data-stu-id="7d0f0-108">Requirement set</span></span>  | <span data-ttu-id="7d0f0-109">Office 2013 for Windows</span><span class="sxs-lookup"><span data-stu-id="7d0f0-109">Office 2013 for Windows</span></span> | <span data-ttu-id="7d0f0-110">Office 2016 for Windows（MSI 安装）</span><span class="sxs-lookup"><span data-stu-id="7d0f0-110">Office 2016 for Windows (MSI Installs)</span></span>   | <span data-ttu-id="7d0f0-111">Office 365 for Windows（C2R 安装）</span><span class="sxs-lookup"><span data-stu-id="7d0f0-111">Office 2016 for Windows (C2R Installs)</span></span>   |  <span data-ttu-id="7d0f0-112">Office 365 for iPad</span><span class="sxs-lookup"><span data-stu-id="7d0f0-112">Office 365 for iPad</span></span>  |  <span data-ttu-id="7d0f0-113">Office 365 for Mac</span><span class="sxs-lookup"><span data-stu-id="7d0f0-113">Office 365 for Mac</span></span>  | <span data-ttu-id="7d0f0-114">Office Online</span><span class="sxs-lookup"><span data-stu-id="7d0f0-114">Office Online</span></span>  |  <span data-ttu-id="7d0f0-115">Office Online Server</span><span class="sxs-lookup"><span data-stu-id="7d0f0-115">Office Online Server</span></span>  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="7d0f0-116">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="7d0f0-116">DialogApi 1.1</span></span>  | <span data-ttu-id="7d0f0-117">生成号 15.0.4855.1000 或更高版本</span><span class="sxs-lookup"><span data-stu-id="7d0f0-117">Build 15.0.4855.1000 or later</span></span> | <span data-ttu-id="7d0f0-118">生成号 16.0.4390.1000 或更高版本</span><span class="sxs-lookup"><span data-stu-id="7d0f0-118">Build 16.0.4390.1000 or later</span></span> | <span data-ttu-id="7d0f0-119">版本 1602（生成号 6741.0000）或更高版本</span><span class="sxs-lookup"><span data-stu-id="7d0f0-119">Version 1602 (Build 6741.0000) or later</span></span> | <span data-ttu-id="7d0f0-120">1.22 或更高版本</span><span class="sxs-lookup"><span data-stu-id="7d0f0-120">1.22 or later</span></span> | <span data-ttu-id="7d0f0-121">15.20 或更高版本</span><span class="sxs-lookup"><span data-stu-id="7d0f0-121">15.20 or later</span></span>| <span data-ttu-id="7d0f0-122">2017 年 1 月</span><span class="sxs-lookup"><span data-stu-id="7d0f0-122">January 2017</span></span> | <span data-ttu-id="7d0f0-123">版本 1608（生成号 7601.6800）或更高版本</span><span class="sxs-lookup"><span data-stu-id="7d0f0-123">Version 1608 (Build 7601.6800) or later</span></span>|

<span data-ttu-id="7d0f0-124">若要详细了解版本、生成号和 Office Online Server，请参阅：</span><span class="sxs-lookup"><span data-stu-id="7d0f0-124">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

- [<span data-ttu-id="7d0f0-125">更新频道发布的 Office 365 客户端版本号和内部版本号</span><span class="sxs-lookup"><span data-stu-id="7d0f0-125">Version and build numbers of update channel releases for Office 365 clients</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="7d0f0-126">使用的是哪一版 Office？</span><span class="sxs-lookup"><span data-stu-id="7d0f0-126">What version of Office am I using?</span></span>](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [<span data-ttu-id="7d0f0-127">在哪里可以找到 Office 365 客户端应用程序的版本号和内部版本号</span><span class="sxs-lookup"><span data-stu-id="7d0f0-127">Where you can find the version and build number for an Office 365 client application</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="7d0f0-128">Office Online Server 概述</span><span class="sxs-lookup"><span data-stu-id="7d0f0-128">Office Online Server overview</span></span>](https://docs.microsoft.com/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="7d0f0-129">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="7d0f0-129">Office common API requirement sets</span></span>

<span data-ttu-id="7d0f0-130">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="7d0f0-130">For information about common API requirement sets, see [Office common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="dialog-api-11"></a><span data-ttu-id="7d0f0-131">Dialog API 1.1</span><span class="sxs-lookup"><span data-stu-id="7d0f0-131">Dialog API 1.1</span></span> 

<span data-ttu-id="7d0f0-132">Dialog API 1.1 是首版 API。</span><span class="sxs-lookup"><span data-stu-id="7d0f0-132">Excel JavaScript API 1.1 is the first version of the API.</span></span> <span data-ttu-id="7d0f0-133">有关 API 的详细信息，请参阅 [Dialog API](/javascript/api/office/office.ui) 参考主题。</span><span class="sxs-lookup"><span data-stu-id="7d0f0-133">For details about the API, see the [getAccessTokenAsync](/javascript/api/office/office.ui) reference topic.</span></span>

## <a name="see-also"></a><span data-ttu-id="7d0f0-134">另请参阅</span><span class="sxs-lookup"><span data-stu-id="7d0f0-134">See also</span></span>

- [<span data-ttu-id="7d0f0-135">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="7d0f0-135">Office versions and requirement sets</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="7d0f0-136">指定 Office 主机和 API 要求</span><span class="sxs-lookup"><span data-stu-id="7d0f0-136">Specify Office hosts and API requirements</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="7d0f0-137">Office 外接程序 XML 清单</span><span class="sxs-lookup"><span data-stu-id="7d0f0-137">Office Add-ins XML manifest</span></span>](https://docs.microsoft.com/office/dev/add-ins/develop/add-in-manifests)
