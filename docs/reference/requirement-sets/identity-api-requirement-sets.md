---
title: Identity API 要求集
description: Office 外接程序的标识 API 要求集信息。
ms.date: 07/30/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: 05805451f17cc70597a61e55d1ecacbb81c383c5
ms.sourcegitcommit: 8fdd7369bfd97a273e222a0404e337ba2b8807b0
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/05/2020
ms.locfileid: "46573215"
---
# <a name="identity-api-requirement-sets"></a><span data-ttu-id="281ba-103">Identity API 要求集</span><span class="sxs-lookup"><span data-stu-id="281ba-103">Identity API requirement sets</span></span>

<span data-ttu-id="281ba-p101">要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 主机是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="281ba-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office host supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="281ba-107">Office 外接程序在多个 Office 版本中运行。</span><span class="sxs-lookup"><span data-stu-id="281ba-107">Office Add-ins run across multiple versions of Office.</span></span> <span data-ttu-id="281ba-108">下表列出了 Identity API 要求集、支持该要求集的 Office 主机应用程序，以及 Office 应用程序的内部版本或版本号。</span><span class="sxs-lookup"><span data-stu-id="281ba-108">The following table lists the Identity API requirement sets, the Office host applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="281ba-109">要求集</span><span class="sxs-lookup"><span data-stu-id="281ba-109">Requirement set</span></span>  | <span data-ttu-id="281ba-110">Windows 上的 Office 2013 或更高版本</span><span class="sxs-lookup"><span data-stu-id="281ba-110">Office 2013 or later on Windows</span></span><br><span data-ttu-id="281ba-111">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="281ba-111">(one-time purchase)</span></span> | <span data-ttu-id="281ba-112">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="281ba-112">Office on Windows</span></span><br><span data-ttu-id="281ba-113">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="281ba-113">(connected to a Microsoft 365 subscription)</span></span> |  <span data-ttu-id="281ba-114">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="281ba-114">Office on iPad</span></span><br><span data-ttu-id="281ba-115">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="281ba-115">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="281ba-116">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="281ba-116">Office on Mac</span></span><br><span data-ttu-id="281ba-117">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="281ba-117">(connected to a Microsoft 365 subscription)</span></span>  | <span data-ttu-id="281ba-118">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="281ba-118">Office on the web</span></span>  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="281ba-119">IdentityAPI 1。3</span><span class="sxs-lookup"><span data-stu-id="281ba-119">IdentityAPI 1.3</span></span>  | <span data-ttu-id="281ba-120">不适用</span><span class="sxs-lookup"><span data-stu-id="281ba-120">N/A</span></span> | <span data-ttu-id="281ba-121">2008 (内部版本 13127.20000) 或更高版本</span><span class="sxs-lookup"><span data-stu-id="281ba-121">2008 (build 13127.20000) or later</span></span> | <span data-ttu-id="281ba-122">即将推出</span><span class="sxs-lookup"><span data-stu-id="281ba-122">Coming soon</span></span> | <span data-ttu-id="281ba-123">16.40 或更高版本</span><span class="sxs-lookup"><span data-stu-id="281ba-123">16.40 or later</span></span> | <span data-ttu-id="281ba-124">2020年8月 \*</span><span class="sxs-lookup"><span data-stu-id="281ba-124">August, 2020\*</span></span> |

> <span data-ttu-id="281ba-125">\*最初，只有从 SharePoint Online 和 OneDrive.com 打开的文档才会在 web 上的 Office 中支持要求集。</span><span class="sxs-lookup"><span data-stu-id="281ba-125">\* Initially, the requirement set is supported in Office on the web only for documents that are opened from SharePoint Online and OneDrive.com.</span></span> <span data-ttu-id="281ba-126">对其他文档的支持稍后将在2020中向 Office 提供。</span><span class="sxs-lookup"><span data-stu-id="281ba-126">Support for other documents will come to Office on the web later in 2020.</span></span>

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="281ba-127">Office 版本和内部版本号</span><span class="sxs-lookup"><span data-stu-id="281ba-127">Office versions and build numbers</span></span>

<span data-ttu-id="281ba-128">若要详细了解版本、内部版本号和 Office Online Server，请参阅：</span><span class="sxs-lookup"><span data-stu-id="281ba-128">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [<span data-ttu-id="281ba-129">Office Online Server 概述</span><span class="sxs-lookup"><span data-stu-id="281ba-129">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="281ba-130">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="281ba-130">Office Common API requirement sets</span></span>

<span data-ttu-id="281ba-131">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="281ba-131">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="identityapi-preview"></a><span data-ttu-id="281ba-132">IdentityAPI 预览</span><span class="sxs-lookup"><span data-stu-id="281ba-132">IdentityAPI Preview</span></span>

<span data-ttu-id="281ba-133">有关此 API 的详细信息，请参阅在[tokenhelper.getaccesstoken 以便](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-)中使用承诺的版本或在[getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-)中使用回调的版本。</span><span class="sxs-lookup"><span data-stu-id="281ba-133">For details about this API, see either the version that uses Promises at [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) or the version that uses callbacks at [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-).</span></span>

## <a name="see-also"></a><span data-ttu-id="281ba-134">另请参阅</span><span class="sxs-lookup"><span data-stu-id="281ba-134">See also</span></span>

- [<span data-ttu-id="281ba-135">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="281ba-135">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="281ba-136">指定 Office 主机和 API 要求</span><span class="sxs-lookup"><span data-stu-id="281ba-136">Specify Office hosts and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="281ba-137">Office 外接程序 XML 清单</span><span class="sxs-lookup"><span data-stu-id="281ba-137">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
