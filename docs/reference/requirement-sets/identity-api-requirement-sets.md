---
title: Identity API 要求集
description: Office 外接程序的标识 API 要求集信息。
ms.date: 01/26/2021
ms.prod: non-product-specific
localization_priority: Normal
ms.openlocfilehash: c662e7a5306692fd75de51acc7cadfd1df3e7406
ms.sourcegitcommit: 85b4839be743059bf155ff44e49d64968444d80a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/31/2021
ms.locfileid: "51471722"
---
# <a name="identity-api-requirement-sets"></a><span data-ttu-id="3a533-103">Identity API 要求集</span><span class="sxs-lookup"><span data-stu-id="3a533-103">Identity API requirement sets</span></span>

<span data-ttu-id="3a533-p101">要求集是指各组已命名的 API 成员。Office 加载项使用清单中指定的要求集或执行运行时检查，以确定 Office 应用程序是否支持加载项所需的 API。有关详细信息，请参阅 [Office 版本和要求集](../../develop/office-versions-and-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="3a533-p101">Requirement sets are named groups of API members. Office Add-ins use requirement sets specified in the manifest or use a runtime check to determine whether an Office application supports APIs that an add-in needs. For more information, see [Office versions and requirement sets](../../develop/office-versions-and-requirement-sets.md).</span></span>

<span data-ttu-id="3a533-107">Office 外接程序在多个 Office 版本中运行。</span><span class="sxs-lookup"><span data-stu-id="3a533-107">Office Add-ins run across multiple versions of Office.</span></span> <span data-ttu-id="3a533-108">下表列出了 Identity API 要求集、支持该要求集的 Office 客户端应用程序，以及 Office 应用程序内部版本或版本号。</span><span class="sxs-lookup"><span data-stu-id="3a533-108">The following table lists the Identity API requirement sets, the Office client applications that support that requirement set, and the build or version numbers for the Office application.</span></span>

|  <span data-ttu-id="3a533-109">要求集</span><span class="sxs-lookup"><span data-stu-id="3a533-109">Requirement set</span></span>  | <span data-ttu-id="3a533-110">Windows 上的 Office 2013 或更高版本</span><span class="sxs-lookup"><span data-stu-id="3a533-110">Office 2013 or later on Windows</span></span><br><span data-ttu-id="3a533-111">（一次性购买）</span><span class="sxs-lookup"><span data-stu-id="3a533-111">(one-time purchase)</span></span> | <span data-ttu-id="3a533-112">Windows 版 Office</span><span class="sxs-lookup"><span data-stu-id="3a533-112">Office on Windows</span></span><br><span data-ttu-id="3a533-113">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3a533-113">(connected to a Microsoft 365 subscription)</span></span> |  <span data-ttu-id="3a533-114">iPad 版 Office</span><span class="sxs-lookup"><span data-stu-id="3a533-114">Office on iPad</span></span><br><span data-ttu-id="3a533-115">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3a533-115">(connected to a Microsoft 365 subscription)</span></span>  |  <span data-ttu-id="3a533-116">Mac 版 Office</span><span class="sxs-lookup"><span data-stu-id="3a533-116">Office on Mac</span></span><br><span data-ttu-id="3a533-117">（关联至 Microsoft 365 订阅）</span><span class="sxs-lookup"><span data-stu-id="3a533-117">(connected to a Microsoft 365 subscription)</span></span>  | <span data-ttu-id="3a533-118">Office 网页版</span><span class="sxs-lookup"><span data-stu-id="3a533-118">Office on the web</span></span>  |
|:-----|-----|:-----|:-----|:-----|:-----|:-----|:-----|:-----|
| <span data-ttu-id="3a533-119">IdentityAPI 1.3</span><span class="sxs-lookup"><span data-stu-id="3a533-119">IdentityAPI 1.3</span></span>  | <span data-ttu-id="3a533-120">不适用</span><span class="sxs-lookup"><span data-stu-id="3a533-120">N/A</span></span> | <span data-ttu-id="3a533-121">2008 (版本 13127.20000) 或更高版本</span><span class="sxs-lookup"><span data-stu-id="3a533-121">2008 (build 13127.20000) or later</span></span> | <span data-ttu-id="3a533-122">即将推出</span><span class="sxs-lookup"><span data-stu-id="3a533-122">Coming soon</span></span> | <span data-ttu-id="3a533-123">16.40 或更高版本</span><span class="sxs-lookup"><span data-stu-id="3a533-123">16.40 or later</span></span> | <span data-ttu-id="3a533-124">Microsoft SharePoint Online 和 OneDrive\*</span><span class="sxs-lookup"><span data-stu-id="3a533-124">Microsoft SharePoint Online and OneDrive\*</span></span> |

<span data-ttu-id="3a533-125">\* 目前，要求集仅在从 Microsoft SharePoint Online 和 OneDrive 打开的文档的 Office 网页版中受支持。</span><span class="sxs-lookup"><span data-stu-id="3a533-125">\* Currently, the requirement set is supported in Office on the web only for documents that are opened from Microsoft SharePoint Online and OneDrive.</span></span>

> [!NOTE]
> <span data-ttu-id="3a533-126">Outlook：若要要求在加载项代码中将 Identity API 设置为 1.3，请通过调用 检查是否受支持 `isSetSupported('IdentityAPI', '1.3')` 。</span><span class="sxs-lookup"><span data-stu-id="3a533-126">Outlook: To require the Identity API set 1.3 in your add-in code, check if it's supported by calling `isSetSupported('IdentityAPI', '1.3')`.</span></span> <span data-ttu-id="3a533-127">不支持在 Outlook 外接程序清单中声明它。</span><span class="sxs-lookup"><span data-stu-id="3a533-127">Declaring it in the Outlook add-in's manifest isn't supported.</span></span> <span data-ttu-id="3a533-128">还可通过检查其不是 `undefined` 来确定该 API 是否受到支持。</span><span class="sxs-lookup"><span data-stu-id="3a533-128">You can also determine if the API is supported by checking that it's not `undefined`.</span></span> <span data-ttu-id="3a533-129">有关详细信息，请参阅 [从后续要求集中使用 API](outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets)。</span><span class="sxs-lookup"><span data-stu-id="3a533-129">For further details, see [Using APIs from later requirement sets](outlook-api-requirement-sets.md#using-apis-from-later-requirement-sets).</span></span>

## <a name="office-versions-and-build-numbers"></a><span data-ttu-id="3a533-130">Office 版本和内部版本号</span><span class="sxs-lookup"><span data-stu-id="3a533-130">Office versions and build numbers</span></span>

<span data-ttu-id="3a533-131">若要详细了解版本、内部版本号和 Office Online Server，请参阅：</span><span class="sxs-lookup"><span data-stu-id="3a533-131">To find out more about versions, build numbers, and Office Online Server, see:</span></span>

[!INCLUDE [Links to get Office versions and how to find Office client version](../../includes/links-get-office-versions-builds.md)]
- [<span data-ttu-id="3a533-132">Office Online Server 概述</span><span class="sxs-lookup"><span data-stu-id="3a533-132">Office Online Server overview</span></span>](/officeonlineserver/office-online-server-overview)

## <a name="office-common-api-requirement-sets"></a><span data-ttu-id="3a533-133">Office 通用 API 要求集</span><span class="sxs-lookup"><span data-stu-id="3a533-133">Office Common API requirement sets</span></span>

<span data-ttu-id="3a533-134">若要了解通用 API 要求集，请参阅 [Office 通用 API 要求集](office-add-in-requirement-sets.md)。</span><span class="sxs-lookup"><span data-stu-id="3a533-134">For information about Common API requirement sets, see [Office Common API requirement sets](office-add-in-requirement-sets.md).</span></span>

## <a name="identityapi-preview"></a><span data-ttu-id="3a533-135">IdentityAPI 预览</span><span class="sxs-lookup"><span data-stu-id="3a533-135">IdentityAPI Preview</span></span>

<span data-ttu-id="3a533-136">有关此 API 的详细信息，请参阅 [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) 处使用 Promises 的版本，或者使用 [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-)的回调的版本。</span><span class="sxs-lookup"><span data-stu-id="3a533-136">For details about this API, see either the version that uses Promises at [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-) or the version that uses callbacks at [getAccessTokenAsync](/javascript/api/office/office.auth#getaccesstokenasync-options--callback-).</span></span>

## <a name="see-also"></a><span data-ttu-id="3a533-137">另请参阅</span><span class="sxs-lookup"><span data-stu-id="3a533-137">See also</span></span>

- [<span data-ttu-id="3a533-138">Office 版本和要求集</span><span class="sxs-lookup"><span data-stu-id="3a533-138">Office versions and requirement sets</span></span>](../../develop/office-versions-and-requirement-sets.md)
- [<span data-ttu-id="3a533-139">指定 Office 应用程序和 API 要求集</span><span class="sxs-lookup"><span data-stu-id="3a533-139">Specify Office applications and API requirements</span></span>](../../develop/specify-office-hosts-and-api-requirements.md)
- [<span data-ttu-id="3a533-140">Office 加载项 XML 清单</span><span class="sxs-lookup"><span data-stu-id="3a533-140">Office Add-ins XML manifest</span></span>](../../develop/add-in-manifests.md)
