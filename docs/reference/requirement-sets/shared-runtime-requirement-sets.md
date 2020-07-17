---
title: 共享运行时要求集
description: 指定支持 SharedRuntime Api 的平台和 Office 主机。
ms.date: 07/10/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 37ab904242a07a5ae7f1f580332f709ac409c6be
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159267"
---
# <a name="shared-runtime-requirement-sets"></a><span data-ttu-id="3cdc1-103">共享运行时要求集</span><span class="sxs-lookup"><span data-stu-id="3cdc1-103">Shared runtime requirement sets</span></span>

<span data-ttu-id="3cdc1-p101">要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="3cdc1-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="3cdc1-107">运行 JavaScript 代码（例如任务窗格、从外接程序命令启动的函数文件和 Excel 自定义函数）的 Office 外接程序的各个部分可以共享单个 JavaScript 运行时。</span><span class="sxs-lookup"><span data-stu-id="3cdc1-107">Parts of an Office Add-in that run JavaScript code, such as task panes, function files launched from add-in commands, and Excel custom functions, can share a single JavaScript runtime.</span></span> <span data-ttu-id="3cdc1-108">这使所有部分都可以共享一组全局变量，共享一组已加载库，并且可以相互通信，而无需通过持久化存储传递邮件。</span><span class="sxs-lookup"><span data-stu-id="3cdc1-108">This enables all the parts to share a set of global variables, to share a set of loaded libraries, and to communicate with each other without having to pass messages through a persisted storage.</span></span>

<span data-ttu-id="3cdc1-109">下表列出了 SharedRuntime 1.1 要求集、支持该要求集的 Office 主机应用程序，以及 Office 应用程序的内部版本号或版本号。</span><span class="sxs-lookup"><span data-stu-id="3cdc1-109">The following table lists the SharedRuntime 1.1 requirement set, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="3cdc1-110">要求集</span><span class="sxs-lookup"><span data-stu-id="3cdc1-110">Requirement set</span></span>  |  <span data-ttu-id="3cdc1-111">Windows 上的 Office 2013 （或更高版本）</span><span class="sxs-lookup"><span data-stu-id="3cdc1-111">Office 2013 (or later) on Windows</span></span><br><span data-ttu-id="3cdc1-112">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3cdc1-112">(one-time purchase)</span></span> | <span data-ttu-id="3cdc1-113">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="3cdc1-113">Office on Windows</span></span><br><span data-ttu-id="3cdc1-114">（连接到 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3cdc1-114">(connected to a Microsoft 365 subscription)</span></span>   |  <span data-ttu-id="3cdc1-115">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="3cdc1-115">Office on iPad</span></span><br><span data-ttu-id="3cdc1-116">（连接到 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3cdc1-116">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="3cdc1-117">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="3cdc1-117">Office on Mac</span></span><br><span data-ttu-id="3cdc1-118">（连接到 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3cdc1-118">(connected to a Microsoft 365 subscription)</span></span>  | <span data-ttu-id="3cdc1-119">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="3cdc1-119">Office on the web</span></span>  | <span data-ttu-id="3cdc1-120">Office Online Server</span><span class="sxs-lookup"><span data-stu-id="3cdc1-120">Office Online Server</span></span> |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="3cdc1-121">SharedRuntime 1。1</span><span class="sxs-lookup"><span data-stu-id="3cdc1-121">SharedRuntime 1.1</span></span>  | <span data-ttu-id="3cdc1-122">不适用</span><span class="sxs-lookup"><span data-stu-id="3cdc1-122">N/A</span></span> | <span data-ttu-id="3cdc1-123">版本2002（内部版本12527.20092）或更高版本</span><span class="sxs-lookup"><span data-stu-id="3cdc1-123">Version 2002 (Build 12527.20092) or later</span></span> | <span data-ttu-id="3cdc1-124">不适用</span><span class="sxs-lookup"><span data-stu-id="3cdc1-124">N/A</span></span> | <span data-ttu-id="3cdc1-125">16.35 或更高版本</span><span class="sxs-lookup"><span data-stu-id="3cdc1-125">16.35 or later</span></span> | <span data-ttu-id="3cdc1-126">2020 年 2 月</span><span class="sxs-lookup"><span data-stu-id="3cdc1-126">February 2020</span></span> | <span data-ttu-id="3cdc1-127">不适用</span><span class="sxs-lookup"><span data-stu-id="3cdc1-127">N/A</span></span> |

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="3cdc1-128">Office 版本和内部版本号</span><span class="sxs-lookup"><span data-stu-id="3cdc1-128">Office versions and build numbers</span></span>

<span data-ttu-id="3cdc1-129">若要详细了解版本、内部版本号和 Office Online Server，请参阅：</span><span class="sxs-lookup"><span data-stu-id="3cdc1-129">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [<span data-ttu-id="3cdc1-130">Office Online Server 概述</span><span class="sxs-lookup"><span data-stu-id="3cdc1-130">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="3cdc1-131">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="3cdc1-131">Office Common API requirement sets</span></span>

<span data-ttu-id="3cdc1-132">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="3cdc1-132">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="3cdc1-133">另请参阅</span><span class="sxs-lookup"><span data-stu-id="3cdc1-133">See also</span></span>

- [<span data-ttu-id="3cdc1-134">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="3cdc1-134">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="3cdc1-135">指定 Office 主机和 API 要求</span><span class="sxs-lookup"><span data-stu-id="3cdc1-135">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="3cdc1-136">Office 外接程序 XML 清单</span><span class="sxs-lookup"><span data-stu-id="3cdc1-136">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
