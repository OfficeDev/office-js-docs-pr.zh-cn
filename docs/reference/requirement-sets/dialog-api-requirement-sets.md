---
title: Dialog API 要求集
description: 了解有关对话框 API 要求集的详细信息
ms.date: 03/11/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: aeae2f140b158f3343c9812db8e9f27ea7608a3e
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/17/2020
ms.locfileid: "42719906"
---
# <a name="dialog-api-requirement-sets"></a><span data-ttu-id="a63b9-103">Dialog API 要求集</span><span class="sxs-lookup"><span data-stu-id="a63b9-103">Dialog API requirement sets</span></span>

<span data-ttu-id="a63b9-p101">要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="a63b9-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="a63b9-p102">Office 外接程序在多个 Office 版本中运行。下表列出了 Dialog API 要求集、支持该要求集的 Office 主机应用程序，以及 Office 应用程序的内部版本或版本号。</span><span class="sxs-lookup"><span data-stu-id="a63b9-p102">Office Add-ins run across multiple versions of Office. The following table lists the Dialog API requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="a63b9-109">要求集</span><span class="sxs-lookup"><span data-stu-id="a63b9-109">Requirement set</span></span>  | <span data-ttu-id="a63b9-110">Windows 版 Office 2013\*</span><span class="sxs-lookup"><span data-stu-id="a63b9-110">Office 2013 on Windows\*</span></span><br><span data-ttu-id="a63b9-111">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="a63b9-111">(one-time purchase)</span></span> | <span data-ttu-id="a63b9-112">Windows 上的 Office 2016 或更高版本\*</span><span class="sxs-lookup"><span data-stu-id="a63b9-112">Office 2016 or later on Windows\*</span></span><br><span data-ttu-id="a63b9-113">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="a63b9-113">(one-time purchase)</span></span>   | <span data-ttu-id="a63b9-114">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="a63b9-114">Office on Windows</span></span><br><span data-ttu-id="a63b9-115">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="a63b9-115">(connected to Office 365 subscription)</span></span> |  <span data-ttu-id="a63b9-116">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="a63b9-116">Office on iPad</span></span><br><span data-ttu-id="a63b9-117">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="a63b9-117">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="a63b9-118">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="a63b9-118">Office on Mac</span></span><br><span data-ttu-id="a63b9-119">（连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="a63b9-119">(connected to Office 365 subscription)</span></span>  | <span data-ttu-id="a63b9-120">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="a63b9-120">Office on the web</span></span>  |  <span data-ttu-id="a63b9-121">Office Online Server</span><span class="sxs-lookup"><span data-stu-id="a63b9-121">Office Online Server</span></span>  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="a63b9-122">DialogApi 1.1</span><span class="sxs-lookup"><span data-stu-id="a63b9-122">DialogApi 1.1</span></span>  | <span data-ttu-id="a63b9-123">生成号 15.0.4855.1000 或更高版本</span><span class="sxs-lookup"><span data-stu-id="a63b9-123">Build 15.0.4855.1000 or later</span></span> | <span data-ttu-id="a63b9-124">生成号 16.0.4390.1000 或更高版本</span><span class="sxs-lookup"><span data-stu-id="a63b9-124">Build 16.0.4390.1000 or later</span></span> | <span data-ttu-id="a63b9-125">版本 1602（生成号 6741.0000）或更高版本</span><span class="sxs-lookup"><span data-stu-id="a63b9-125">Version 1602 (Build 6741.0000) or later</span></span> | <span data-ttu-id="a63b9-126">1.22 或更高版本</span><span class="sxs-lookup"><span data-stu-id="a63b9-126">1.22 or later</span></span> | <span data-ttu-id="a63b9-127">15.20 或更高版本</span><span class="sxs-lookup"><span data-stu-id="a63b9-127">15.20 or later</span></span>| <span data-ttu-id="a63b9-128">2017 年 1 月</span><span class="sxs-lookup"><span data-stu-id="a63b9-128">January 2017</span></span> | <span data-ttu-id="a63b9-129">版本 1608（内部版本 7601.6800）或更高版本</span><span class="sxs-lookup"><span data-stu-id="a63b9-129">Version 1608 (Build 7601.6800) or later</span></span>|

><span data-ttu-id="a63b9-130">\*一次性购买 Office 的用户可能未接受所有修补和更新。</span><span class="sxs-lookup"><span data-stu-id="a63b9-130">\* Users of the one-time purchase Office may not have accepted all patches and updates.</span></span> <span data-ttu-id="a63b9-131">如果是这样，即使在用户的计算机上未安装支持 DialogApi 所需的更新的 Dll，Office 用来在 UI 中报告其版本的 DLL 可能也会大于此处列出的版本。</span><span class="sxs-lookup"><span data-stu-id="a63b9-131">If so, the DLL that Office uses to report its version in the UI may be greater than the versions listed here even if the updated DLLs needed to support DialogApi have not be installed on the user's computer.</span></span> <span data-ttu-id="a63b9-132">若要确保安装了所需的修补程序，用户必须转到 Office 更新列表（[office 2013 列表](/officeupdates/msp-files-office-2013)或[office 2016 列表](/officeupdates/msp-files-office-2016)），搜索**osfclient-x**，并安装列出的修补程序。</span><span class="sxs-lookup"><span data-stu-id="a63b9-132">To ensure that the needed patch is installed, the user must go to the Office update list ([Office 2013 list](/officeupdates/msp-files-office-2013) or [Office 2016 list](/officeupdates/msp-files-office-2016)), search for **osfclient-x-none**, and install the listed patch.</span></span>

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="a63b9-133">Office 版本和内部版本号</span><span class="sxs-lookup"><span data-stu-id="a63b9-133">Office versions and build numbers</span></span>

<span data-ttu-id="a63b9-134">若要详细了解版本、内部版本号和 Office Online Server，请参阅：</span><span class="sxs-lookup"><span data-stu-id="a63b9-134">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [<span data-ttu-id="a63b9-135">Office Online Server 概述</span><span class="sxs-lookup"><span data-stu-id="a63b9-135">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="a63b9-136">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="a63b9-136">Office Common API requirement sets</span></span>

<span data-ttu-id="a63b9-137">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="a63b9-137">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="dialog-api-11"></a><span data-ttu-id="a63b9-138">Dialog API 1.1</span><span class="sxs-lookup"><span data-stu-id="a63b9-138">Dialog API 1.1</span></span>

<span data-ttu-id="a63b9-139">Dialog API 1.1 是首版 API。</span><span class="sxs-lookup"><span data-stu-id="a63b9-139">The Dialog API 1.1 is the first version of the API.</span></span> <span data-ttu-id="a63b9-140">有关 API 的详细信息，请参阅[对话框 API](/javascript/api/office/office.ui)参考主题。</span><span class="sxs-lookup"><span data-stu-id="a63b9-140">For details about the API, see the [Dialog API](/javascript/api/office/office.ui) reference topic.</span></span>

## <a name="see-also"></a><span data-ttu-id="a63b9-141">另请参阅</span><span class="sxs-lookup"><span data-stu-id="a63b9-141">See also</span></span>

- [<span data-ttu-id="a63b9-142">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="a63b9-142">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="a63b9-143">指定 Office 主机和 API 要求</span><span class="sxs-lookup"><span data-stu-id="a63b9-143">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="a63b9-144">Office 外接程序 XML 清单</span><span class="sxs-lookup"><span data-stu-id="a63b9-144">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
