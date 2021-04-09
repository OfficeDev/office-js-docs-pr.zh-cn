---
title: 共享运行时要求集
description: 指定支持 SharedRuntime API 的平台和 Office 应用程序。
ms.date: 04/08/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 8d0db6e129aaf7a4aa2967e7a1341d6db1188359
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652221"
---
# <a name="shared-runtime-requirement-sets"></a><span data-ttu-id="84390-103">共享运行时要求集</span><span class="sxs-lookup"><span data-stu-id="84390-103">Shared runtime requirement sets</span></span>

<span data-ttu-id="84390-p101">要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="84390-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="84390-107">运行 JavaScript 代码的 Office 外接程序的某些部分（如任务窗格、从外接程序命令启动的函数文件和 Excel 自定义函数）可以共享单个 JavaScript 运行时。</span><span class="sxs-lookup"><span data-stu-id="84390-107">Parts of an Office Add-in that run JavaScript code, such as task panes, function files launched from add-in commands, and Excel custom functions, can share a single JavaScript runtime.</span></span> <span data-ttu-id="84390-108">这允许所有部件共享一组全局变量、共享一组加载的库以及相互通信，而无需通过持久存储传递邮件。</span><span class="sxs-lookup"><span data-stu-id="84390-108">This enables all the parts to share a set of global variables, to share a set of loaded libraries, and to communicate with each other without having to pass messages through a persisted storage.</span></span> <span data-ttu-id="84390-109">有关详细信息，请参阅将 [Office 加载项配置为使用共享的 JavaScript 运行时](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)。</span><span class="sxs-lookup"><span data-stu-id="84390-109">For more information, see [Configure your Office Add-in to use a shared JavaScript runtime](../../develop/configure-your-add-in-to-use-a-shared-runtime.md).</span></span>

<span data-ttu-id="84390-110">下表列出了 SharedRuntime 1.1 要求集、支持该要求集的 Office 客户端应用程序，以及 Office 应用程序内部版本或版本号。</span><span class="sxs-lookup"><span data-stu-id="84390-110">The following table lists the SharedRuntime 1.1 requirement set, the Office client applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="84390-111">要求集</span><span class="sxs-lookup"><span data-stu-id="84390-111">Requirement set</span></span>  |  <span data-ttu-id="84390-112">Windows 版 Office 2013 (或) 更高版本</span><span class="sxs-lookup"><span data-stu-id="84390-112">Office 2013 (or later) on Windows</span></span><br><span data-ttu-id="84390-113">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="84390-113">(one-time purchase)</span></span> | <span data-ttu-id="84390-114">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="84390-114">Office on Windows</span></span><br><span data-ttu-id="84390-115">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="84390-115">(connected to a Microsoft 365 subscription)</span></span>   |  <span data-ttu-id="84390-116">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="84390-116">Office on iPad</span></span><br><span data-ttu-id="84390-117">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="84390-117">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="84390-118">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="84390-118">Office on Mac</span></span><br><span data-ttu-id="84390-119">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="84390-119">(connected to a Microsoft 365 subscription)</span></span>  | <span data-ttu-id="84390-120">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="84390-120">Office on the web</span></span>  | <span data-ttu-id="84390-121">Office Online Server</span><span class="sxs-lookup"><span data-stu-id="84390-121">Office Online Server</span></span> |
|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="84390-122">SharedRuntime 1.1</span><span class="sxs-lookup"><span data-stu-id="84390-122">SharedRuntime 1.1</span></span>  | <span data-ttu-id="84390-123">不适用</span><span class="sxs-lookup"><span data-stu-id="84390-123">N/A</span></span> | <span data-ttu-id="84390-124">版本 2002 (内部版本 12527.20092) 或更高版本</span><span class="sxs-lookup"><span data-stu-id="84390-124">Version 2002 (Build 12527.20092) or later</span></span> | <span data-ttu-id="84390-125">不适用</span><span class="sxs-lookup"><span data-stu-id="84390-125">N/A</span></span> | <span data-ttu-id="84390-126">16.35 或更高版本</span><span class="sxs-lookup"><span data-stu-id="84390-126">16.35 or later</span></span> | <span data-ttu-id="84390-127">2020 年 2 月</span><span class="sxs-lookup"><span data-stu-id="84390-127">February 2020</span></span> | <span data-ttu-id="84390-128">不适用</span><span class="sxs-lookup"><span data-stu-id="84390-128">N/A</span></span> |

> [!IMPORTANT]
> <span data-ttu-id="84390-129">共享的 JavaScript 运行时要求集仅在以下平台上可用。</span><span class="sxs-lookup"><span data-stu-id="84390-129">The shared JavaScript runtime requirement set is only available on the following platforms.</span></span>
>
> - <span data-ttu-id="84390-130">Excel 网页版、Windows 和 Mac。</span><span class="sxs-lookup"><span data-stu-id="84390-130">Excel on the web, Windows, and Mac.</span></span>
> - <span data-ttu-id="84390-131">Windows 版 PowerPoint（内部版本 13218.10000 或更高版本）。</span><span class="sxs-lookup"><span data-stu-id="84390-131">PowerPoint on Windows (build 13218.10000 or later).</span></span> <span data-ttu-id="84390-132">适用于 PowerPoint 的共享 JavaScript 运行时当前处于预览阶段并可能会发生更改。</span><span class="sxs-lookup"><span data-stu-id="84390-132">The shared JavaScript runtime for PowerPoint is currently in preview and subject to change.</span></span> <span data-ttu-id="84390-133">不支持在生产环境中使用。</span><span class="sxs-lookup"><span data-stu-id="84390-133">It is not supported for use in production environments.</span></span> <span data-ttu-id="84390-134">要获取最新版本，你需要[加入 Office 预览体验计划](https://insider.office.com/join)。</span><span class="sxs-lookup"><span data-stu-id="84390-134">To get the latest build you need to [join Office Insider](https://insider.office.com/join).</span></span> <span data-ttu-id="84390-135">试用预览版功能的好方法是使用 Microsoft 365 订阅。</span><span class="sxs-lookup"><span data-stu-id="84390-135">A good way to try out preview features is by using a Microsoft 365 subscription.</span></span> <span data-ttu-id="84390-136">如果还没有 Microsoft 365 订阅，可以通过加入[Microsoft 365 开发人员计划](https://developer.microsoft.com/office/dev-program)获取一个订阅。</span><span class="sxs-lookup"><span data-stu-id="84390-136">If you don't already have a Microsoft 365 subscription, you can get one by joining the [Microsoft 365 developer program](https://developer.microsoft.com/office/dev-program).</span></span>
>
> <span data-ttu-id="84390-137">目前，iPad 或一次性购买版本的 Office 2019 或更早版本不支持共享 JavaScript 运行时。</span><span class="sxs-lookup"><span data-stu-id="84390-137">At this time, the shared JavaScript runtime is not supported on iPad or in one-time purchase versions of Office 2019 or earlier.</span></span>

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="84390-138">Office 版本和内部版本号</span><span class="sxs-lookup"><span data-stu-id="84390-138">Office versions and build numbers</span></span>

<span data-ttu-id="84390-139">若要详细了解版本、内部版本号和 Office Online Server，请参阅：</span><span class="sxs-lookup"><span data-stu-id="84390-139">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [<span data-ttu-id="84390-140">Office Online Server 概述</span><span class="sxs-lookup"><span data-stu-id="84390-140">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="84390-141">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="84390-141">Office Common API requirement sets</span></span>

<span data-ttu-id="84390-142">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="84390-142">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="84390-143">另请参阅</span><span class="sxs-lookup"><span data-stu-id="84390-143">See also</span></span>

- [<span data-ttu-id="84390-144">将 Office 加载项配置为使用共享 JavaScript 运行时</span><span class="sxs-lookup"><span data-stu-id="84390-144">Configure your Office Add-in to use a shared JavaScript runtime</span></span>](../../develop/configure-your-add-in-to-use-a-shared-runtime.md)
- [<span data-ttu-id="84390-145">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="84390-145">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="84390-146">指定 Office 应用程序和 API 要求集</span><span class="sxs-lookup"><span data-stu-id="84390-146">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="84390-147">Office 加载项 XML 清单</span><span class="sxs-lookup"><span data-stu-id="84390-147">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
