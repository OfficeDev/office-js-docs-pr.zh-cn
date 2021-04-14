---
title: 打开浏览器窗口要求集
description: 指定支持 openBrowserWindow API 的 Office 平台和内部版本。
ms.date: 04/09/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: dd15136b350d42ec49187e436142aaecbfe70f40
ms.sourcegitcommit: 841bcad3c6c5139fd0953707c0be73ce890fa463
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/13/2021
ms.locfileid: "51687431"
---
# <a name="open-browser-window-api-requirement-sets"></a><span data-ttu-id="67994-103">打开浏览器窗口 API 要求集</span><span class="sxs-lookup"><span data-stu-id="67994-103">Open Browser Window API requirement sets</span></span>

<span data-ttu-id="67994-p101">要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="67994-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="67994-107">OpenBrowserWindow API 集使加载项能够打开浏览器，以完成无法在加载项本身的沙盒 Webview 控件中始终完成的任务;例如，在 Microsoft Edge 提供 Webview 控件时下载 PDF 文件。</span><span class="sxs-lookup"><span data-stu-id="67994-107">The OpenBrowserWindow API set enables add-ins to open a browser to accomplish tasks that cannot always be done in the sandboxed webview control within the add-in itself; for example, downloading a PDF file when the webview control is provided by Microsoft Edge.</span></span>

<span data-ttu-id="67994-108">Office 外接程序在多个 Office 版本中运行。</span><span class="sxs-lookup"><span data-stu-id="67994-108">Office Add-ins run across multiple versions of Office.</span></span> <span data-ttu-id="67994-109">下表列出了 OpenBrowserWindow API 要求集、支持该要求集的 Office 主机应用程序，以及 Office 应用程序内部版本或版本号。</span><span class="sxs-lookup"><span data-stu-id="67994-109">The following table lists the OpenBrowserWindow API requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="67994-110">要求集</span><span class="sxs-lookup"><span data-stu-id="67994-110">Requirement set</span></span>  | <span data-ttu-id="67994-111">Windows 版 Office 2013 或更高版本</span><span class="sxs-lookup"><span data-stu-id="67994-111">Office 2013 on Windows or later</span></span><br><span data-ttu-id="67994-112">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="67994-112">(one-time purchase)</span></span> | <span data-ttu-id="67994-113">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="67994-113">Office on Windows</span></span><br><span data-ttu-id="67994-114">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="67994-114">(connected to Microsoft 365 subscription)</span></span> |  <span data-ttu-id="67994-115">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="67994-115">Office on iPad</span></span><br><span data-ttu-id="67994-116">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="67994-116">(connected to Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="67994-117">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="67994-117">Office on Mac</span></span><br><span data-ttu-id="67994-118">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="67994-118">(connected to Microsoft 365 subscription)</span></span>  | <span data-ttu-id="67994-119">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="67994-119">Office on the web</span></span>  |  <span data-ttu-id="67994-120">Office Online Server</span><span class="sxs-lookup"><span data-stu-id="67994-120">Office Online Server</span></span>  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="67994-121">OpenBrowserWindowApi 1.1</span><span class="sxs-lookup"><span data-stu-id="67994-121">OpenBrowserWindowApi 1.1</span></span>  | <span data-ttu-id="67994-122">不适用</span><span class="sxs-lookup"><span data-stu-id="67994-122">N/A</span></span> | <span data-ttu-id="67994-123">版本 1810 (内部版本 16.0.11001.20074) 或更高版本</span><span class="sxs-lookup"><span data-stu-id="67994-123">Version 1810 (Build 16.0.11001.20074) or later</span></span> | <span data-ttu-id="67994-124">16.0.0.0 或更高版本</span><span class="sxs-lookup"><span data-stu-id="67994-124">16.0.0.0 or later</span></span> | <span data-ttu-id="67994-125">16.0.0.0 或更高版本</span><span class="sxs-lookup"><span data-stu-id="67994-125">16.0.0.0 or later</span></span> | <span data-ttu-id="67994-126">不适用</span><span class="sxs-lookup"><span data-stu-id="67994-126">N/A</span></span> | <span data-ttu-id="67994-127">不适用</span><span class="sxs-lookup"><span data-stu-id="67994-127">N/A</span></span>|

> [!NOTE]
> <span data-ttu-id="67994-128">OpenBrowserWindowApi 要求集仅按如下方式提供：</span><span class="sxs-lookup"><span data-stu-id="67994-128">The OpenBrowserWindowApi requirement set is only available as follows:</span></span>
>
> - <span data-ttu-id="67994-129">Excel、PowerPoint、Word：Windows、Mac、iPad</span><span class="sxs-lookup"><span data-stu-id="67994-129">Excel, PowerPoint, Word: Windows, Mac, iPad</span></span>
> - <span data-ttu-id="67994-130">Outlook：Windows、Mac</span><span class="sxs-lookup"><span data-stu-id="67994-130">Outlook: Windows, Mac</span></span>

<span data-ttu-id="67994-131">若要详细了解版本、内部版本号和 Office Online Server，请参阅：</span><span class="sxs-lookup"><span data-stu-id="67994-131">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

- [<span data-ttu-id="67994-132">Microsoft 365 应用更新频道版本的版本号和内部版本号</span><span class="sxs-lookup"><span data-stu-id="67994-132">Version and build numbers of update channel releases for Microsoft 365 Apps</span></span>](/officeupdates/update-history-microsoft365-apps-by-date)
- [<span data-ttu-id="67994-133">使用的是哪一版 Office？</span><span class="sxs-lookup"><span data-stu-id="67994-133">What version of Office am I using?</span></span>](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [<span data-ttu-id="67994-134">在哪里可以找到 Office 客户端应用程序的版本号和内部版本号</span><span class="sxs-lookup"><span data-stu-id="67994-134">Where you can find the version and build number for an Office client application</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="67994-135">Office Online Server 概述</span><span class="sxs-lookup"><span data-stu-id="67994-135">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="67994-136">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="67994-136">Office Common API requirement sets</span></span>

<span data-ttu-id="67994-137">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="67994-137">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="openbrowserwindowapi-11"></a><span data-ttu-id="67994-138">OpenBrowserWindowApi 1.1</span><span class="sxs-lookup"><span data-stu-id="67994-138">OpenBrowserWindowApi 1.1</span></span>

<span data-ttu-id="67994-139">OpenBrowserWindowApi 1.1 是 API 的第一个版本。</span><span class="sxs-lookup"><span data-stu-id="67994-139">The OpenBrowserWindowApi 1.1 is the first version of the API.</span></span> <span data-ttu-id="67994-140">有关 API 的详细信息，请参阅 [Office.context.ui](/javascript/api/office/office.context#ui) 参考主题。</span><span class="sxs-lookup"><span data-stu-id="67994-140">For details about the API, see the [Office.context.ui](/javascript/api/office/office.context#ui) reference topic.</span></span>

## <a name="see-also"></a><span data-ttu-id="67994-141">另请参阅</span><span class="sxs-lookup"><span data-stu-id="67994-141">See also</span></span>

- [<span data-ttu-id="67994-142">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="67994-142">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="67994-143">指定 Office 主机和 API 要求</span><span class="sxs-lookup"><span data-stu-id="67994-143">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="67994-144">Office 加载项 XML 清单</span><span class="sxs-lookup"><span data-stu-id="67994-144">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
