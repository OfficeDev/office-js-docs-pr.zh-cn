---
title: 向 Azure AD v2.0 端点注册使用 SSO 的 Office 外接程序
description: ''
ms.date: 04/10/2018
ms.openlocfilehash: 95b690e21bddf7f2754cc308c8b771e629bbc630
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437253"
---
# <a name="register-an-office-add-in-that-uses-sso-with-the-azure-ad-v20-endpoint"></a><span data-ttu-id="37ee9-102">向 Azure AD v2.0 端点注册使用 SSO 的 Office 外接程序</span><span class="sxs-lookup"><span data-stu-id="37ee9-102">Details are at: Register an Office Add-in that uses SSO with the Azure AD v2.0 endpoint.</span></span>

<span data-ttu-id="37ee9-103">本文介绍了如何向 Azure AD v2.0 端点注册 Office 外接程序。</span><span class="sxs-lookup"><span data-stu-id="37ee9-103">This article explains how to register an Office Add-in with the Azure AD v2.0 endpoint.</span></span> <span data-ttu-id="37ee9-104">开始开发时，需要注册外接程序。</span><span class="sxs-lookup"><span data-stu-id="37ee9-104">You need to register the add-in when you begin developing it.</span></span> <span data-ttu-id="37ee9-105">进行测试或生产时，可以为外接 程序的开发、测试和生产版本更改现有注册或创建单独的注册。</span><span class="sxs-lookup"><span data-stu-id="37ee9-105">When you progress to testing or production, you can change the existing registration or create separate registrations for development, testing, and production versions of the add-in.</span></span> 

<span data-ttu-id="37ee9-106">下表列出了执行此过程所需的信息以及说明中出现的相应占位符。</span><span class="sxs-lookup"><span data-stu-id="37ee9-106">The following table itemizes the information that you need to carry out this procedure and the corresponding placeholders that appear in the instructions.</span></span> 

|<span data-ttu-id="37ee9-107">信息</span><span class="sxs-lookup"><span data-stu-id="37ee9-107">Information</span></span>  |<span data-ttu-id="37ee9-108">示例</span><span class="sxs-lookup"><span data-stu-id="37ee9-108">Examples</span></span>  |<span data-ttu-id="37ee9-109">占位符</span><span class="sxs-lookup"><span data-stu-id="37ee9-109">Placeholder</span></span>  |
|---------|---------|---------|
|<span data-ttu-id="37ee9-110">外接程序的人类可读名称。</span><span class="sxs-lookup"><span data-stu-id="37ee9-110">A human readable name for the add-in.</span></span> <span data-ttu-id="37ee9-111">（建议使用唯一名称，但并非强制性要求。）</span><span class="sxs-lookup"><span data-stu-id="37ee9-111">(Uniqueness recommended, but not required.)</span></span>    |`Contoso Marketing Excel Add-in (Prod)`        |<span data-ttu-id="37ee9-112">**$ADD-IN-NAME$**</span><span class="sxs-lookup"><span data-stu-id="37ee9-112">**$ADD-IN-NAME$**</span></span>         |
|<span data-ttu-id="37ee9-113">外接 程序的完全限定的域名（协议除外）。</span><span class="sxs-lookup"><span data-stu-id="37ee9-113">The fully qualified domain name (except for protocol) of the add-in.</span></span> <span data-ttu-id="37ee9-114">*必须使用你所拥有的域名。*</span><span class="sxs-lookup"><span data-stu-id="37ee9-114">*You must use a domain that you own.*</span></span> <span data-ttu-id="37ee9-115">出于这个原因，你不能使用某些众所周知的领域，如 `azurewebsites.net` 或者 `cloudapp.net`。</span><span class="sxs-lookup"><span data-stu-id="37ee9-115">For this reason, you cannot use certain well-known domains such as `azurewebsites.net` or `cloudapp.net`.</span></span>   |<span data-ttu-id="37ee9-116">`localhost:6789`, `addins.contoso.com`</span><span class="sxs-lookup"><span data-stu-id="37ee9-116"></span></span>         |<span data-ttu-id="37ee9-117">**$FQDN-WITHOUT-PROTOCOL$**</span><span class="sxs-lookup"><span data-stu-id="37ee9-117">**$FQDN-WITHOUT-PROTOCOL$**</span></span>         |
|<span data-ttu-id="37ee9-118">外接程序所需的 AAD 和 Microsoft Graph 权限。</span><span class="sxs-lookup"><span data-stu-id="37ee9-118">The permissions to AAD and Microsoft Graph that your add-in needs.</span></span> <span data-ttu-id="37ee9-119">（始终需要 `profile`。）</span><span class="sxs-lookup"><span data-stu-id="37ee9-119">(`profile` is always required.)</span></span>    |<span data-ttu-id="37ee9-120">`profile`, `Files.Read.All`</span><span class="sxs-lookup"><span data-stu-id="37ee9-120"></span></span>         |<span data-ttu-id="37ee9-121">无</span><span class="sxs-lookup"><span data-stu-id="37ee9-121">N/A</span></span>         |

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]