---
title: 向 Azure AD v2.0 端点注册使用 SSO 的 Office 加载项。
description: 了解如何使用 Azure AD v2.0 终结点注册 Office 外接程序。
ms.date: 04/10/2019
localization_priority: Normal
ms.openlocfilehash: 8bcd72bd6f2d56c5f97d2d4f153d6791d111452e
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609372"
---
# <a name="register-an-office-add-in-that-uses-sso-with-the-azure-ad-v20-endpoint"></a><span data-ttu-id="081d5-103">向 Azure AD v2.0 端点注册使用 SSO 的 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="081d5-103">Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint</span></span>

<span data-ttu-id="081d5-104">本文介绍如何向 Azure AD v2.0 端点注册 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="081d5-104">This article explains how to register an Office Add-in with the Azure AD v2.0 endpoint.</span></span> <span data-ttu-id="081d5-105">开始开发时，需要注册加载项。</span><span class="sxs-lookup"><span data-stu-id="081d5-105">You need to register the add-in when you begin developing it.</span></span> <span data-ttu-id="081d5-106">在进行测试或生产时，可以更改现有注册或为加载项的开发、测试和生产版本创建单独的注册。</span><span class="sxs-lookup"><span data-stu-id="081d5-106">When you progress to testing or production, you can change the existing registration or create separate registrations for development, testing, and production versions of the add-in.</span></span>

<span data-ttu-id="081d5-107">下表列出了执行此过程所需的信息以及说明中显示的相应占位符。</span><span class="sxs-lookup"><span data-stu-id="081d5-107">The following table itemizes the information that you need to carry out this procedure and the corresponding placeholders that appear in the instructions.</span></span>

|<span data-ttu-id="081d5-108">信息</span><span class="sxs-lookup"><span data-stu-id="081d5-108">Information</span></span>  |<span data-ttu-id="081d5-109">示例</span><span class="sxs-lookup"><span data-stu-id="081d5-109">Examples</span></span>  |<span data-ttu-id="081d5-110">占位符</span><span class="sxs-lookup"><span data-stu-id="081d5-110">Placeholder</span></span>  |
|---------|---------|---------|
|<span data-ttu-id="081d5-111">加载项的人类可读名称。</span><span class="sxs-lookup"><span data-stu-id="081d5-111">A human readable name for the add-in.</span></span> <span data-ttu-id="081d5-112">（建议使用唯一名称，但不是必需的。）</span><span class="sxs-lookup"><span data-stu-id="081d5-112">(Uniqueness recommended, but not required.)</span></span>|`Contoso Marketing Excel Add-in (Prod)`|<span data-ttu-id="081d5-113">**$ADD-IN-NAME$**</span><span class="sxs-lookup"><span data-stu-id="081d5-113">**$ADD-IN-NAME$**</span></span>|
|<span data-ttu-id="081d5-114">加载项的完全限定域名（协议除外）。</span><span class="sxs-lookup"><span data-stu-id="081d5-114">The fully qualified domain name (except for protocol) of the add-in.</span></span> <span data-ttu-id="081d5-115">*必须使用自己的域*。</span><span class="sxs-lookup"><span data-stu-id="081d5-115">*You must use a domain that you own.*</span></span> <span data-ttu-id="081d5-116">正因如此，不能使用某些知名域名，例如 `azurewebsites.net` 或 `cloudapp.net`。</span><span class="sxs-lookup"><span data-stu-id="081d5-116">For this reason, you cannot use certain well-known domains such as `azurewebsites.net` or `cloudapp.net`.</span></span> <span data-ttu-id="081d5-117">域必须相同，包括任何子域，如加载项清单的 `<Resources>` 部分中的 URL 中所使用的那样。</span><span class="sxs-lookup"><span data-stu-id="081d5-117">The domain must be the same, including any subdomains, as is used in the URLs in the `<Resources>` section of the add-in's manifest.</span></span>|<span data-ttu-id="081d5-118">`localhost:6789`, `addins.contoso.com`</span><span class="sxs-lookup"><span data-stu-id="081d5-118">`localhost:6789`, `addins.contoso.com`</span></span>|<span data-ttu-id="081d5-119">**$FQDN-WITHOUT-PROTOCOL$**</span><span class="sxs-lookup"><span data-stu-id="081d5-119">**$FQDN-WITHOUT-PROTOCOL$**</span></span>|
|<span data-ttu-id="081d5-120">加载项所需的 AAD 和 Microsoft Graph 权限。</span><span class="sxs-lookup"><span data-stu-id="081d5-120">The permissions to AAD and Microsoft Graph that your add-in needs.</span></span> <span data-ttu-id="081d5-121">（`profile` 始终是必需的。）</span><span class="sxs-lookup"><span data-stu-id="081d5-121">(`profile` is always required.)</span></span>|<span data-ttu-id="081d5-122">`profile`, `Files.Read.All`</span><span class="sxs-lookup"><span data-stu-id="081d5-122">`profile`, `Files.Read.All`</span></span>|<span data-ttu-id="081d5-123">不适用</span><span class="sxs-lookup"><span data-stu-id="081d5-123">N/A</span></span>|

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]
