---
title: 共享运行时要求集
description: 指定支持 SharedRuntime Api 的平台和 Office 主机。
ms.date: 02/11/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: dbb9d908154da074eaff6901c778adea168504a9
ms.sourcegitcommit: 7464eac3b54a6a6b65e27549a3ad603af6ee1011
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/27/2020
ms.locfileid: "42315878"
---
# <a name="shared-runtime-requirement-sets"></a><span data-ttu-id="df54e-103">共享运行时要求集</span><span class="sxs-lookup"><span data-stu-id="df54e-103">Shared runtime requirement sets</span></span>

<span data-ttu-id="df54e-p101">要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](/office/dev/add-ins/develop/office-versions-and-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="df54e-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](/office/dev/add-ins/develop/office-versions-and-requirement-sets).</span></span>

<span data-ttu-id="df54e-107">运行 JavaScript 代码（例如任务窗格、从外接程序命令启动的函数文件和 Excel 自定义函数）的 Office 外接程序的各个部分可以共享单个 JavaScript 运行时。</span><span class="sxs-lookup"><span data-stu-id="df54e-107">Parts of an Office Add-in that run JavaScript code, such as task panes, function files launched from add-in commands, and Excel custom functions, can share a single JavaScript runtime.</span></span> <span data-ttu-id="df54e-108">这使所有部分都可以共享一组全局变量，共享一组已加载库，并且可以相互通信，而无需通过持久化存储传递邮件。</span><span class="sxs-lookup"><span data-stu-id="df54e-108">This enables all the parts to share a set of global variables, to share a set of loaded libraries, and to communicate with each other without having to pass messages through a persisted storage.</span></span>

<span data-ttu-id="df54e-109">下表列出了 SharedRuntime 1.1 要求集、支持该要求集的 Office 主机应用程序，以及 Office 应用程序的内部版本号或版本号。</span><span class="sxs-lookup"><span data-stu-id="df54e-109">The following table lists the SharedRuntime 1.1 requirement set, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="df54e-110">要求集</span><span class="sxs-lookup"><span data-stu-id="df54e-110">Requirement set</span></span>  |  <span data-ttu-id="df54e-111">Windows 上的 Office 2013 （或更高版本）</span><span class="sxs-lookup"><span data-stu-id="df54e-111">Office 2013 (or later) on Windows</span></span><br><span data-ttu-id="df54e-112">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="df54e-112">(one-time purchase)</span></span> | <span data-ttu-id="df54e-113">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="df54e-113">Office on Windows</span></span><br><span data-ttu-id="df54e-114">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="df54e-114">(connected to Office 365 subscription)</span></span>   |  <span data-ttu-id="df54e-115">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="df54e-115">Office on iPad</span></span><br><span data-ttu-id="df54e-116">（已连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="df54e-116">(connected to Office 365 subscription)</span></span>  |  <span data-ttu-id="df54e-117">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="df54e-117">Office on Mac</span></span><br><span data-ttu-id="df54e-118">（连接到 Office 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="df54e-118">(connected to Office 365 subscription)</span></span>  | <span data-ttu-id="df54e-119">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="df54e-119">Office on the web</span></span>  | <span data-ttu-id="df54e-120">Office Online Server</span><span class="sxs-lookup"><span data-stu-id="df54e-120">Office Online Server</span></span> |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="df54e-121">SharedRuntime 1。1</span><span class="sxs-lookup"><span data-stu-id="df54e-121">SharedRuntime 1.1</span></span>  | <span data-ttu-id="df54e-122">不适用</span><span class="sxs-lookup"><span data-stu-id="df54e-122">N/A</span></span> | <span data-ttu-id="df54e-123">版本2002（内部版本12527.20092）或更高版本</span><span class="sxs-lookup"><span data-stu-id="df54e-123">Version 2002 (Build 12527.20092) or later</span></span> | <span data-ttu-id="df54e-124">不适用</span><span class="sxs-lookup"><span data-stu-id="df54e-124">N/A</span></span> | <span data-ttu-id="df54e-125">16.35 或更高版本</span><span class="sxs-lookup"><span data-stu-id="df54e-125">16.35 or later</span></span> | <span data-ttu-id="df54e-126">2020 年 2 月</span><span class="sxs-lookup"><span data-stu-id="df54e-126">February 2020</span></span> | <span data-ttu-id="df54e-127">不适用</span><span class="sxs-lookup"><span data-stu-id="df54e-127">N/A</span></span> |

<span data-ttu-id="df54e-128">若要详细了解版本、内部版本号和 Office Online Server，请参阅：</span><span class="sxs-lookup"><span data-stu-id="df54e-128">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

- [<span data-ttu-id="df54e-129">更新频道发布的 Office 365 客户端版本号和内部版本号</span><span class="sxs-lookup"><span data-stu-id="df54e-129">Version and build numbers of update channel releases for Office 365 clients</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="df54e-130">使用的是哪一版 Office？</span><span class="sxs-lookup"><span data-stu-id="df54e-130">What version of Office am I using?</span></span>](https://support.office.com/article/What-version-of-Office-am-I-using-932788b8-a3ce-44bf-bb09-e334518b8b19)
- [<span data-ttu-id="df54e-131">在哪里可以找到 Office 365 客户端应用程序的版本号和内部版本号</span><span class="sxs-lookup"><span data-stu-id="df54e-131">Where you can find the version and build number for an Office 365 client application</span></span>](https://support.office.com/article/version-and-build-numbers-of-update-channel-releases-ae942449-1fca-4484-898b-a933ea23def7)
- [<span data-ttu-id="df54e-132">Office Online Server 概述</span><span class="sxs-lookup"><span data-stu-id="df54e-132">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="df54e-133">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="df54e-133">Office Common API requirement sets</span></span>

<span data-ttu-id="df54e-134">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="df54e-134">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="df54e-135">另请参阅</span><span class="sxs-lookup"><span data-stu-id="df54e-135">See also</span></span>

- [<span data-ttu-id="df54e-136">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="df54e-136">Office versions and requirement sets</span></span>](/office/dev/add-ins/develop/office-versions-and-requirement-sets)
- [<span data-ttu-id="df54e-137">指定 Office 主机和 API 要求</span><span class="sxs-lookup"><span data-stu-id="df54e-137">Specify Office hosts and API requirements</span></span>](/office/dev/add-ins/develop/specify-office-hosts-and-api-requirements)
- [<span data-ttu-id="df54e-138">Office 外接程序 XML 清单</span><span class="sxs-lookup"><span data-stu-id="df54e-138">Office Add-ins XML manifest</span></span>](/office/dev/add-ins/develop/add-in-manifests)
