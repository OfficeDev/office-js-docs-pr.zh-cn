---
title: Outlook 加载项要求
description: 必须满足服务器和客户端的多个要求，才能正常加载和运行 Outlook 加载项。
ms.date: 07/07/2020
localization_priority: Priority
ms.openlocfilehash: 353c03fc0cdfe83c5f775df09dfb7c6b23cca191
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/28/2020
ms.locfileid: "47294001"
---
# <a name="outlook-add-in-requirements"></a><span data-ttu-id="ee8b8-103">Outlook 加载项要求</span><span class="sxs-lookup"><span data-stu-id="ee8b8-103">Outlook add-in requirements</span></span>

<span data-ttu-id="ee8b8-104">必须满足服务器和客户端的多个要求，才能正常加载和运行 Outlook 加载项。</span><span class="sxs-lookup"><span data-stu-id="ee8b8-104">For Outlook add-ins to load and function properly, there are a number of requirements for both the servers and the clients.</span></span>

## <a name="client-requirements"></a><span data-ttu-id="ee8b8-105">客户端要求</span><span class="sxs-lookup"><span data-stu-id="ee8b8-105">Client requirements</span></span>

- <span data-ttu-id="ee8b8-106">客户端必须是一个受支持的 Outlook 加载项应用程序。下列客户端支持加载项：</span><span class="sxs-lookup"><span data-stu-id="ee8b8-106">The client must be one of the supported applications for Outlook add-ins. The following clients support add-ins:</span></span>

   - <span data-ttu-id="ee8b8-107">Windows 版 Outlook 2013 或更高版本</span><span class="sxs-lookup"><span data-stu-id="ee8b8-107">Outlook 2013 or later on Windows</span></span>
   - <span data-ttu-id="ee8b8-108">Mac 版 Outlook 2016 或更高版本</span><span class="sxs-lookup"><span data-stu-id="ee8b8-108">Outlook 2016 or later on Mac</span></span>
   - <span data-ttu-id="ee8b8-109">iOS 版 Outlook</span><span class="sxs-lookup"><span data-stu-id="ee8b8-109">Outlook on iOS</span></span>
   - <span data-ttu-id="ee8b8-110">Android 版 Outlook</span><span class="sxs-lookup"><span data-stu-id="ee8b8-110">Outlook on Android</span></span>
   - <span data-ttu-id="ee8b8-111">适用于 Exchange 2016 或更高版本和 Office 365 的 Outlook 网页版</span><span class="sxs-lookup"><span data-stu-id="ee8b8-111">Outlook on the web for Exchange 2016 or later and Office 365</span></span>
   - <span data-ttu-id="ee8b8-112">适用于 Exchange 2013 的 Outlook 网页版</span><span class="sxs-lookup"><span data-stu-id="ee8b8-112">Outlook on the web for Exchange 2013</span></span>
   - <span data-ttu-id="ee8b8-113">Outlook.com</span><span class="sxs-lookup"><span data-stu-id="ee8b8-113">Outlook.com</span></span>

- <span data-ttu-id="ee8b8-p101">必须使用直接连接将客户端连接到 Exchange 服务器或 Microsoft 365。在配置客户端时，用户必须选择 **Exchange**、**Office 365** 或 **Outlook.com** 帐户类型。如果将客户端配置为使用 POP3 或 IMAP 连接，将不会加载加载项。</span><span class="sxs-lookup"><span data-stu-id="ee8b8-p101">The client must be connected to an Exchange server or Microsoft 365 using a direct connection. When configuring the client, the user must choose an **Exchange**, **Office 365**, or **Outlook.com** account type. If the client is configured to connect with POP3 or IMAP, add-ins will not load.</span></span>

## <a name="mail-server-requirements"></a><span data-ttu-id="ee8b8-117">邮件服务器要求</span><span class="sxs-lookup"><span data-stu-id="ee8b8-117">Mail server requirements</span></span>

<span data-ttu-id="ee8b8-p102">如果用户已连接到 Microsoft 365 或 Outlook.com，则已经满足了所有邮件服务器要求。但是，对于连接到 Exchange Server 本地安装的用户，适用以下要求。</span><span class="sxs-lookup"><span data-stu-id="ee8b8-p102">If the user is connected to Microsoft 365 or Outlook.com, mail server requirements are all taken care of already. However, for users connected to on-premises installations of Exchange Server, the following requirements apply.</span></span>

- <span data-ttu-id="ee8b8-120">服务器必须是 Exchange 2013 或更高版本。</span><span class="sxs-lookup"><span data-stu-id="ee8b8-120">The server must be Exchange 2013 or later.</span></span>
- <span data-ttu-id="ee8b8-121">必须启用 Exchange Web 服务 (EWS)，并向 Internet 公开此服务。</span><span class="sxs-lookup"><span data-stu-id="ee8b8-121">Exchange Web Services (EWS) must be enabled and must be exposed to the Internet.</span></span> <span data-ttu-id="ee8b8-122">许多加载项要求，必须启用 EWS 才能正常运行。</span><span class="sxs-lookup"><span data-stu-id="ee8b8-122">Many add-ins require EWS to function properly.</span></span>
- <span data-ttu-id="ee8b8-123">服务器必须有有效身份验证证书，才能颁发有效标识令牌。</span><span class="sxs-lookup"><span data-stu-id="ee8b8-123">The server must have a valid authentication certificate in order for the server to issue valid identity tokens.</span></span> <span data-ttu-id="ee8b8-124">新安装的 Exchange Server 包含默认身份验证证书。</span><span class="sxs-lookup"><span data-stu-id="ee8b8-124">New installations of Exchange Server include a default authentication certificate.</span></span> <span data-ttu-id="ee8b8-125">有关详细信息，请参阅 [Exchange 2016 中的数字证书和加密](/Exchange/architecture/client-access/certificates)和 [Set-AuthConfig](/powershell/module/exchange/organization/Set-AuthConfig)。</span><span class="sxs-lookup"><span data-stu-id="ee8b8-125">For more information, see [Digital certificates and encryption in Exchange 2016](/Exchange/architecture/client-access/certificates) and [Set-AuthConfig](/powershell/module/exchange/organization/Set-AuthConfig).</span></span>
- <span data-ttu-id="ee8b8-126">客户端访问服务器必须能够与 AppSource 通信，才能从 [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=a35323d5-0e3d-4cc0-ba44-57537d74aae8&omexanonuid=581941df-1c6f-4eda-89e7-651af8aeaeb2) 获取加载项。</span><span class="sxs-lookup"><span data-stu-id="ee8b8-126">To access add-ins from [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=a35323d5-0e3d-4cc0-ba44-57537d74aae8&omexanonuid=581941df-1c6f-4eda-89e7-651af8aeaeb2), the client access servers must be able to communicate with AppSource.</span></span>

## <a name="add-in-server-requirements"></a><span data-ttu-id="ee8b8-127">加载项服务器要求</span><span class="sxs-lookup"><span data-stu-id="ee8b8-127">Add-in server requirements</span></span>

<span data-ttu-id="ee8b8-p105">可在任意需要的 Web 服务器平台上托管外接程序文件（HTML、JavaScript 等）。唯一的要求是，必须将服务器配置为使用 HTTPS，并且 SSL 证书必须受客户端信任。</span><span class="sxs-lookup"><span data-stu-id="ee8b8-p105">Add-in files (HTML, JavaScript, etc.) can be hosted on any web server platform desired. The only requirement is that the server must be configured to use HTTPS, and the SSL certificate must be trusted by the client.</span></span>

## <a name="see-also"></a><span data-ttu-id="ee8b8-130">另请参阅</span><span class="sxs-lookup"><span data-stu-id="ee8b8-130">See also</span></span>

- [<span data-ttu-id="ee8b8-131">Office 加载项的运行要求</span><span class="sxs-lookup"><span data-stu-id="ee8b8-131">Requirements for running Office Add-ins</span></span>](../concepts/requirements-for-running-office-add-ins.md)
- [<span data-ttu-id="ee8b8-132">Office 客户端应用程序和 Office 加载项的平台可用性（Outlook 部分）</span><span class="sxs-lookup"><span data-stu-id="ee8b8-132">Office client application and platform availability for Office Add-ins (Outlook section)</span></span>](../overview/office-add-in-availability.md#outlook)
- [<span data-ttu-id="ee8b8-133">Outlook JavaScript API 要求集支持</span><span class="sxs-lookup"><span data-stu-id="ee8b8-133">Outlook JavaScript API requirement set support</span></span>](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)
