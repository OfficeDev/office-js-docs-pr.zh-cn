---
title: 向 Azure AD v2.0 端点注册使用 SSO 的 Office 加载项。
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: 7b9c0dbcdf8a892ffcb810972c4d3674acbc31f7
ms.sourcegitcommit: 2ac7d64bb2db75ace516a604866850fce5cb2174
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/14/2018
ms.locfileid: "26298535"
---
# <a name="register-an-office-add-in-that-uses-sso-with-the-azure-ad-v20-endpoint"></a><span data-ttu-id="80373-102">向 Azure AD v2.0 端点注册使用 SSO 的 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="80373-102">Details are at: Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint.</span></span>

<span data-ttu-id="80373-103">本文介绍如何向 Azure AD v2.0 端点注册 Office 加载项。</span><span class="sxs-lookup"><span data-stu-id="80373-103">This article explains how to register an Office Add-in with the Azure AD v2.0 endpoint.</span></span> <span data-ttu-id="80373-104">开始开发时，需要注册加载项。</span><span class="sxs-lookup"><span data-stu-id="80373-104">You need to register the add-in when you begin developing it.</span></span> <span data-ttu-id="80373-105">在进行测试或生产时，可以更改现有注册或为加载项的开发、测试和生产版本创建单独的注册。</span><span class="sxs-lookup"><span data-stu-id="80373-105">When you progress to testing or production, you can change the existing registration or create separate registrations for development, testing, and production versions of the add-in.</span></span>

<span data-ttu-id="80373-106">下表列出了执行此过程所需的信息以及说明中显示的相应占位符。</span><span class="sxs-lookup"><span data-stu-id="80373-106">The following table itemizes the information that you need to carry out this procedure and the corresponding placeholders that appear in the instructions.</span></span> 

|<span data-ttu-id="80373-107">信息</span><span class="sxs-lookup"><span data-stu-id="80373-107">Information</span></span>  |<span data-ttu-id="80373-108">示例</span><span class="sxs-lookup"><span data-stu-id="80373-108">Examples</span></span>  |<span data-ttu-id="80373-109">占位符</span><span class="sxs-lookup"><span data-stu-id="80373-109">Placeholder</span></span>  |
|---------|---------|---------|
|<span data-ttu-id="80373-110">加载项的人类可读名称。</span><span class="sxs-lookup"><span data-stu-id="80373-110">A human readable name for the add-in.</span></span> <span data-ttu-id="80373-111">（建议使用唯一名称，但不是必需的。）</span><span class="sxs-lookup"><span data-stu-id="80373-111">(Uniqueness recommended, but not required.)</span></span>    |`Contoso Marketing Excel Add-in (Prod)`        |<span data-ttu-id="80373-112">**$ADD-IN-NAME$**</span><span class="sxs-lookup"><span data-stu-id="80373-112">**$ADD-IN-NAME$**</span></span>         |
|<span data-ttu-id="80373-113">加载项的完全限定域名（协议除外）。</span><span class="sxs-lookup"><span data-stu-id="80373-113">The fully qualified domain name (except for protocol) of the add-in.</span></span> <span data-ttu-id="80373-114">*必须使用自己的域*。</span><span class="sxs-lookup"><span data-stu-id="80373-114">*You must use a domain that you own.*</span></span> <span data-ttu-id="80373-115">正因如此，不能使用某些知名域名，例如 `azurewebsites.net` 或 `cloudapp.net`。</span><span class="sxs-lookup"><span data-stu-id="80373-115">For this reason, you cannot use certain well-known domains such as `azurewebsites.net` or `cloudapp.net`.</span></span> <span data-ttu-id="80373-116">域必须相同，包括任何子域，如加载项清单的 `<Resources>` 部分中的 URL 中所使用的那样。</span><span class="sxs-lookup"><span data-stu-id="80373-116">The domain must be the same, including any subdomains, as is used in the URLs in the `<Resources>` section of the add-in's manifest.</span></span>  |<span data-ttu-id="80373-117">`localhost:6789`, `addins.contoso.com`</span><span class="sxs-lookup"><span data-stu-id="80373-117"></span></span>         |<span data-ttu-id="80373-118">**$FQDN-WITHOUT-PROTOCOL$**</span><span class="sxs-lookup"><span data-stu-id="80373-118">**$FQDN-WITHOUT-PROTOCOL$**</span></span>         |
|<span data-ttu-id="80373-119">加载项所需的 AAD 和 Microsoft Graph 权限。</span><span class="sxs-lookup"><span data-stu-id="80373-119">The permissions to AAD and Microsoft Graph that your add-in needs.</span></span> <span data-ttu-id="80373-120">（`profile` 始终是必需的。）</span><span class="sxs-lookup"><span data-stu-id="80373-120">(`profile` is always required.)</span></span>    |<span data-ttu-id="80373-121">`profile`, `Files.Read.All`</span><span class="sxs-lookup"><span data-stu-id="80373-121"></span></span>         |<span data-ttu-id="80373-122">不适用</span><span class="sxs-lookup"><span data-stu-id="80373-122">N/A</span></span>         |

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]