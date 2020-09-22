---
title: 打开浏览器窗口要求集
description: 指定哪些 Office 平台和生成支持 openBrowserWindow API。
ms.date: 09/16/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 8bc26525bf64ed87d46d85cd1248f79696d67f2b
ms.sourcegitcommit: 4a03d8b3f676ee2d91114813cb81bce5da3c8d6b
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/22/2020
ms.locfileid: "48175505"
---
# <a name="open-browser-window-api-requirement-sets"></a><span data-ttu-id="d3376-103">打开浏览器窗口 API 要求集</span><span class="sxs-lookup"><span data-stu-id="d3376-103">Open Browser Window API requirement sets</span></span>

<span data-ttu-id="d3376-p101">要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="d3376-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="d3376-107">OpenBrowserWindow API 集使外接程序能够打开浏览器以完成在外接程序本身的沙盒视图控件中无法始终完成的任务;例如，在 Microsoft Edge 提供 web 视图控件时下载 PDF 文件。</span><span class="sxs-lookup"><span data-stu-id="d3376-107">The OpenBrowserWindow API set enables add-ins to open a browser to accomplish tasks that cannot always be done in the sandboxed webview control within the add-in itself; for example, downloading a PDF file when the webview control is provided by Microsoft Edge.</span></span>

<span data-ttu-id="d3376-108">Office 外接程序在多个 Office 版本中运行。</span><span class="sxs-lookup"><span data-stu-id="d3376-108">Office Add-ins run across multiple versions of Office.</span></span> <span data-ttu-id="d3376-109">下表列出了 OpenBrowserWindow API 要求集、支持该要求集的 Office 主机应用程序，以及 Office 应用程序的内部版本号或版本号。</span><span class="sxs-lookup"><span data-stu-id="d3376-109">The following table lists the OpenBrowserWindow API requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="d3376-110">要求集</span><span class="sxs-lookup"><span data-stu-id="d3376-110">Requirement set</span></span>  | <span data-ttu-id="d3376-111">Windows 或更高版本上的 Office 2013</span><span class="sxs-lookup"><span data-stu-id="d3376-111">Office 2013 on Windows or later</span></span><br><span data-ttu-id="d3376-112">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="d3376-112">(one-time purchase)</span></span> | <span data-ttu-id="d3376-113">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="d3376-113">Office on Windows</span></span><br><span data-ttu-id="d3376-114">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="d3376-114">(connected to Office 365 subscription)</span></span> |  <span data-ttu-id="d3376-115">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="d3376-115">Office on iPad</span></span><br><span data-ttu-id="d3376-116">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="d3376-116">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="d3376-117">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="d3376-117">Office on Mac</span></span><br><span data-ttu-id="d3376-118">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="d3376-118">(connected to Office 365 subscription)</span></span>  | <span data-ttu-id="d3376-119">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="d3376-119">Office on the web</span></span>  |  <span data-ttu-id="d3376-120">Office Online Server</span><span class="sxs-lookup"><span data-stu-id="d3376-120">Office Online Server</span></span>  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="d3376-121">OpenBrowserWindowApi 1。1</span><span class="sxs-lookup"><span data-stu-id="d3376-121">OpenBrowserWindowApi 1.1</span></span>  | <span data-ttu-id="d3376-122">无</span><span class="sxs-lookup"><span data-stu-id="d3376-122">N/A</span></span> | <span data-ttu-id="d3376-123">版本 1810 (内部版本 16.0.11001.20074) 或更高版本</span><span class="sxs-lookup"><span data-stu-id="d3376-123">Version 1810 (Build 16.0.11001.20074) or later</span></span> | <span data-ttu-id="d3376-124">16.0.0.0 或更高版本</span><span class="sxs-lookup"><span data-stu-id="d3376-124">16.0.0.0 or later</span></span> | <span data-ttu-id="d3376-125">16.0.0.0 或更高版本</span><span class="sxs-lookup"><span data-stu-id="d3376-125">16.0.0.0 or later</span></span> | <span data-ttu-id="d3376-126">不适用</span><span class="sxs-lookup"><span data-stu-id="d3376-126">N/A</span></span> | <span data-ttu-id="d3376-127">不适用</span><span class="sxs-lookup"><span data-stu-id="d3376-127">N/A</span></span>|

<span data-ttu-id="d3376-128">若要详细了解版本、内部版本号和 Office Online Server，请参阅：</span><span class="sxs-lookup"><span data-stu-id="d3376-128">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

- [<span data-ttu-id="d3376-129">更新频道发布的 Office 365 客户端版本号和内部版本号</span><span class="sxs-lookup"><span data-stu-id="d3376-129">Version and build numbers of update channel releases for Office 365 clients</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="d3376-130">使用的是哪一版 Office？</span><span class="sxs-lookup"><span data-stu-id="d3376-130">What version of Office am I using?</span></span>](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [<span data-ttu-id="d3376-131">在哪里可以找到 Office 365 客户端应用程序的版本号和内部版本号</span><span class="sxs-lookup"><span data-stu-id="d3376-131">Where you can find the version and build number for an Office 365 client application</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="d3376-132">Office Online Server 概述</span><span class="sxs-lookup"><span data-stu-id="d3376-132">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="d3376-133">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="d3376-133">Office Common API requirement sets</span></span>

<span data-ttu-id="d3376-134">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="d3376-134">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="openbrowserwindowapi-11"></a><span data-ttu-id="d3376-135">OpenBrowserWindowApi 1。1</span><span class="sxs-lookup"><span data-stu-id="d3376-135">OpenBrowserWindowApi 1.1</span></span>

<span data-ttu-id="d3376-136">OpenBrowserWindowApi 1.1 是 API 的第一个版本。</span><span class="sxs-lookup"><span data-stu-id="d3376-136">The OpenBrowserWindowApi 1.1 is the first version of the API.</span></span> <span data-ttu-id="d3376-137">有关 API 的详细信息，请参阅 " [Office. ui](/javascript/api/office/office.context#ui) 参考" 主题。</span><span class="sxs-lookup"><span data-stu-id="d3376-137">For details about the API, see the [Office.context.ui](/javascript/api/office/office.context#ui) reference topic.</span></span>

## <a name="see-also"></a><span data-ttu-id="d3376-138">另请参阅</span><span class="sxs-lookup"><span data-stu-id="d3376-138">See also</span></span>

- [<span data-ttu-id="d3376-139">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="d3376-139">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="d3376-140">指定 Office 主机和 API 要求</span><span class="sxs-lookup"><span data-stu-id="d3376-140">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="d3376-141">Office 加载项 XML 清单</span><span class="sxs-lookup"><span data-stu-id="d3376-141">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
