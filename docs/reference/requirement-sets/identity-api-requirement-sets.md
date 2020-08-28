---
title: Identity API 要求集
description: Office 外接程序的标识 API 要求集信息。
ms.date: 07/30/2020
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: c2c6ea449cef08248a9ba79051b7c0c5f9baa600
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293539"
---
# <a name="identity-api-requirement-sets"></a><span data-ttu-id="ebf58-103">Identity API 要求集</span><span class="sxs-lookup"><span data-stu-id="ebf58-103">Identity API requirement sets</span></span>

<span data-ttu-id="ebf58-104">要求集是指各组已命名的 API 成员。</span><span class="sxs-lookup"><span data-stu-id="ebf58-104">Requirement sets are named groups of API members.</span></span> <span data-ttu-id="ebf58-105">Office 外接程序使用清单中指定的要求集或使用运行时检查来确定 Office 应用程序是否支持加载项所需的 Api。</span><span class="sxs-lookup"><span data-stu-id="ebf58-105">Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs.</span></span> <span data-ttu-id="ebf58-106">有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="ebf58-106">For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="ebf58-107">Office 外接程序在多个 Office 版本中运行。</span><span class="sxs-lookup"><span data-stu-id="ebf58-107">Office Add-ins run across multiple versions of Office.</span></span> <span data-ttu-id="ebf58-108">下表列出了标识 API 要求集、支持该要求集的 Office 客户端应用程序，以及 Office 应用程序的内部版本号或版本号。</span><span class="sxs-lookup"><span data-stu-id="ebf58-108">The following table lists the Identity API requirement sets, the Office client applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="ebf58-109">要求集</span><span class="sxs-lookup"><span data-stu-id="ebf58-109">Requirement set</span></span>  | <span data-ttu-id="ebf58-110">Windows 上的 Office 2013 或更高版本</span><span class="sxs-lookup"><span data-stu-id="ebf58-110">Office 2013 or later on Windows</span></span><br><span data-ttu-id="ebf58-111">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="ebf58-111">(one-time purchase)</span></span> | <span data-ttu-id="ebf58-112">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="ebf58-112">Office on Windows</span></span><br><span data-ttu-id="ebf58-113">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ebf58-113">(connected to a Microsoft 365 subscription)</span></span> |  <span data-ttu-id="ebf58-114">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="ebf58-114">Office on iPad</span></span><br><span data-ttu-id="ebf58-115">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ebf58-115">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="ebf58-116">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="ebf58-116">Office on Mac</span></span><br><span data-ttu-id="ebf58-117">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="ebf58-117">(connected to a Microsoft 365 subscription)</span></span>  | <span data-ttu-id="ebf58-118">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="ebf58-118">Office on the web</span></span>  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="ebf58-119">IdentityAPI 1。3</span><span class="sxs-lookup"><span data-stu-id="ebf58-119">IdentityAPI 1.3</span></span>  | <span data-ttu-id="ebf58-120">无</span><span class="sxs-lookup"><span data-stu-id="ebf58-120">N/A</span></span> | <span data-ttu-id="ebf58-121">2008 (内部版本 13127.20000) 或更高版本</span><span class="sxs-lookup"><span data-stu-id="ebf58-121">2008 (build 13127.20000) or later</span></span> | <span data-ttu-id="ebf58-122">即将推出</span><span class="sxs-lookup"><span data-stu-id="ebf58-122">Coming soon</span></span> | <span data-ttu-id="ebf58-123">16.40 或更高版本</span><span class="sxs-lookup"><span data-stu-id="ebf58-123">16.40 or later</span></span> | <span data-ttu-id="ebf58-124">2020年8月 \*</span><span class="sxs-lookup"><span data-stu-id="ebf58-124">August, 2020\*</span></span> |

> <span data-ttu-id="ebf58-125">\* 最初，只有从 SharePoint Online 和 OneDrive.com 打开的文档才会在 web 上的 Office 中支持要求集。</span><span class="sxs-lookup"><span data-stu-id="ebf58-125">\* Initially, the requirement set is supported in Office on the web only for documents that are opened from SharePoint Online and OneDrive.com.</span></span> <span data-ttu-id="ebf58-126">对其他文档的支持稍后将在2020中向 Office 提供。</span><span class="sxs-lookup"><span data-stu-id="ebf58-126">Support for other documents will come to Office on the web later in 2020.</span></span>

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="ebf58-127">Office 版本和内部版本号</span><span class="sxs-lookup"><span data-stu-id="ebf58-127">Office versions and build numbers</span></span>

<span data-ttu-id="ebf58-128">若要详细了解版本、内部版本号和 Office Online Server，请参阅：</span><span class="sxs-lookup"><span data-stu-id="ebf58-128">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [<span data-ttu-id="ebf58-129">Office Online Server 概述</span><span class="sxs-lookup"><span data-stu-id="ebf58-129">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="ebf58-130">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="ebf58-130">Office Common API requirement sets</span></span>

<span data-ttu-id="ebf58-131">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="ebf58-131">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="identityapi-preview"></a><span data-ttu-id="ebf58-132">IdentityAPI 预览</span><span class="sxs-lookup"><span data-stu-id="ebf58-132">IdentityAPI Preview</span></span>

<span data-ttu-id="ebf58-133">有关此 API 的详细信息，请参阅在 [tokenhelper.getaccesstoken 以便](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) 中使用承诺的版本或在 [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-)中使用回调的版本。</span><span class="sxs-lookup"><span data-stu-id="ebf58-133">For details about this API, see either the version that uses Promises at [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) or the version that uses callbacks at [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-).</span></span>

## <a name="see-also"></a><span data-ttu-id="ebf58-134">另请参阅</span><span class="sxs-lookup"><span data-stu-id="ebf58-134">See also</span></span>

- [<span data-ttu-id="ebf58-135">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="ebf58-135">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="ebf58-136">指定 Office 应用程序和 API 要求</span><span class="sxs-lookup"><span data-stu-id="ebf58-136">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="ebf58-137">Office 外接程序 XML 清单</span><span class="sxs-lookup"><span data-stu-id="ebf58-137">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
