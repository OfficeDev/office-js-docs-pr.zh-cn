---
title: Outlook 加载项要求
description: 必须满足服务器和客户端的多个要求，才能正常加载和运行 Outlook 加载项。
ms.date: 07/07/2020
localization_priority: Priority
ms.openlocfilehash: 700e0efd2ab2655de61d37d42038fa2c15a99cb4
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/10/2020
ms.locfileid: "45093992"
---
# <a name="outlook-add-in-requirements"></a><span data-ttu-id="7ecae-103">Outlook 加载项要求</span><span class="sxs-lookup"><span data-stu-id="7ecae-103">Outlook add-in requirements</span></span>

<span data-ttu-id="7ecae-104">必须满足服务器和客户端的多个要求，才能正常加载和运行 Outlook 加载项。</span><span class="sxs-lookup"><span data-stu-id="7ecae-104">For Outlook add-ins to load and function properly, there are a number of requirements for both the servers and the clients.</span></span>

## <a name="client-requirements"></a><span data-ttu-id="7ecae-105">客户端要求</span><span class="sxs-lookup"><span data-stu-id="7ecae-105">Client requirements</span></span>

- <span data-ttu-id="7ecae-106">客户端必须是一个受 Outlook 加载项支持的主机。下列客户端支持加载项：</span><span class="sxs-lookup"><span data-stu-id="7ecae-106">The client must be one of the supported hosts for Outlook add-ins. The following clients support add-ins:</span></span>

   - <span data-ttu-id="7ecae-107">Windows 版 Outlook 2013 或更高版本</span><span class="sxs-lookup"><span data-stu-id="7ecae-107">Outlook 2013 or later on Windows</span></span>
   - <span data-ttu-id="7ecae-108">Mac 版 Outlook 2016 或更高版本</span><span class="sxs-lookup"><span data-stu-id="7ecae-108">Outlook 2016 or later on Mac</span></span>
   - <span data-ttu-id="7ecae-109">iOS 版 Outlook</span><span class="sxs-lookup"><span data-stu-id="7ecae-109">Outlook on iOS</span></span>
   - <span data-ttu-id="7ecae-110">Android 版 Outlook</span><span class="sxs-lookup"><span data-stu-id="7ecae-110">Outlook on Android</span></span>
   - <span data-ttu-id="7ecae-111">适用于 Exchange 2016 或更高版本和 Office 365 的 Outlook 网页版</span><span class="sxs-lookup"><span data-stu-id="7ecae-111">Outlook on the web for Exchange 2016 or later and Office 365</span></span>
   - <span data-ttu-id="7ecae-112">适用于 Exchange 2013 的 Outlook 网页版</span><span class="sxs-lookup"><span data-stu-id="7ecae-112">Outlook on the web for Exchange 2013</span></span>
   - <span data-ttu-id="7ecae-113">Outlook.com</span><span class="sxs-lookup"><span data-stu-id="7ecae-113">Outlook.com</span></span>

- <span data-ttu-id="7ecae-114">The client must be connected to an Exchange server or Microsoft 365 using a direct connection.</span><span class="sxs-lookup"><span data-stu-id="7ecae-114">The client must be connected to an Exchange server or Microsoft 365 using a direct connection.</span></span> <span data-ttu-id="7ecae-115">When configuring the client, the user must choose an **Exchange**, **Office 365**, or **Outlook.com** account type.</span><span class="sxs-lookup"><span data-stu-id="7ecae-115">When configuring the client, the user must choose an **Exchange**, **Office 365**, or **Outlook.com** account type.</span></span> <span data-ttu-id="7ecae-116">If the client is configured to connect with POP3 or IMAP, add-ins will not load.</span><span class="sxs-lookup"><span data-stu-id="7ecae-116">If the client is configured to connect with POP3 or IMAP, add-ins will not load.</span></span>

## <a name="mail-server-requirements"></a><span data-ttu-id="7ecae-117">邮件服务器要求</span><span class="sxs-lookup"><span data-stu-id="7ecae-117">Mail server requirements</span></span>

<span data-ttu-id="7ecae-118">If the user is connected to Microsoft 365 or Outlook.com, mail server requirements are all taken care of already.</span><span class="sxs-lookup"><span data-stu-id="7ecae-118">If the user is connected to Microsoft 365 or Outlook.com, mail server requirements are all taken care of already.</span></span> <span data-ttu-id="7ecae-119">However, for users connected to on-premises installations of Exchange Server, the following requirements apply.</span><span class="sxs-lookup"><span data-stu-id="7ecae-119">However, for users connected to on-premises installations of Exchange Server, the following requirements apply.</span></span>

- <span data-ttu-id="7ecae-120">服务器必须是 Exchange 2013 或更高版本。</span><span class="sxs-lookup"><span data-stu-id="7ecae-120">The server must be Exchange 2013 or later.</span></span>
- <span data-ttu-id="7ecae-121">必须启用 Exchange Web 服务 (EWS)，并向 Internet 公开此服务。</span><span class="sxs-lookup"><span data-stu-id="7ecae-121">Exchange Web Services (EWS) must be enabled and must be exposed to the Internet.</span></span> <span data-ttu-id="7ecae-122">许多加载项要求，必须启用 EWS 才能正常运行。</span><span class="sxs-lookup"><span data-stu-id="7ecae-122">Many add-ins require EWS to function properly.</span></span>
- <span data-ttu-id="7ecae-123">服务器必须有有效身份验证证书，才能颁发有效标识令牌。</span><span class="sxs-lookup"><span data-stu-id="7ecae-123">The server must have a valid authentication certificate in order for the server to issue valid identity tokens.</span></span> <span data-ttu-id="7ecae-124">新安装的 Exchange Server 包含默认身份验证证书。</span><span class="sxs-lookup"><span data-stu-id="7ecae-124">New installations of Exchange Server include a default authentication certificate.</span></span> <span data-ttu-id="7ecae-125">有关详细信息，请参阅 [Exchange 2016 中的数字证书和加密](/Exchange/architecture/client-access/certificates)和 [Set-AuthConfig](/powershell/module/exchange/organization/Set-AuthConfig)。</span><span class="sxs-lookup"><span data-stu-id="7ecae-125">For more information, see [Digital certificates and encryption in Exchange 2016](/Exchange/architecture/client-access/certificates) and [Set-AuthConfig](/powershell/module/exchange/organization/Set-AuthConfig).</span></span>
- <span data-ttu-id="7ecae-126">客户端访问服务器必须能够与 AppSource 通信，才能从 [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=a35323d5-0e3d-4cc0-ba44-57537d74aae8&omexanonuid=581941df-1c6f-4eda-89e7-651af8aeaeb2) 获取加载项。</span><span class="sxs-lookup"><span data-stu-id="7ecae-126">To access add-ins from [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=a35323d5-0e3d-4cc0-ba44-57537d74aae8&omexanonuid=581941df-1c6f-4eda-89e7-651af8aeaeb2), the client access servers must be able to communicate with AppSource.</span></span>

## <a name="add-in-server-requirements"></a><span data-ttu-id="7ecae-127">加载项服务器要求</span><span class="sxs-lookup"><span data-stu-id="7ecae-127">Add-in server requirements</span></span>

<span data-ttu-id="7ecae-128">Add-in files (HTML, JavaScript, etc.) can be hosted on any web server platform desired.</span><span class="sxs-lookup"><span data-stu-id="7ecae-128">Add-in files (HTML, JavaScript, etc.) can be hosted on any web server platform desired.</span></span> <span data-ttu-id="7ecae-129">The only requirement is that the server must be configured to use HTTPS, and the SSL certificate must be trusted by the client.</span><span class="sxs-lookup"><span data-stu-id="7ecae-129">The only requirement is that the server must be configured to use HTTPS, and the SSL certificate must be trusted by the client.</span></span>

## <a name="see-also"></a><span data-ttu-id="7ecae-130">另请参阅</span><span class="sxs-lookup"><span data-stu-id="7ecae-130">See also</span></span>

- [<span data-ttu-id="7ecae-131">Office 加载项的运行要求</span><span class="sxs-lookup"><span data-stu-id="7ecae-131">Requirements for running Office Add-ins</span></span>](../concepts/requirements-for-running-office-add-ins.md)
- [<span data-ttu-id="7ecae-132">Office 加载项主机和平台可用性（Outlook 部分）</span><span class="sxs-lookup"><span data-stu-id="7ecae-132">Office Add-in host and platform availability (Outlook section)</span></span>](../overview/office-add-in-availability.md#outlook)
- [<span data-ttu-id="7ecae-133">Outlook JavaScript API 要求集支持</span><span class="sxs-lookup"><span data-stu-id="7ecae-133">Outlook JavaScript API requirement set support</span></span>](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)
