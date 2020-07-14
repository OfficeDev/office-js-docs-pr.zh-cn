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
# <a name="outlook-add-in-requirements"></a>Outlook 加载项要求

必须满足服务器和客户端的多个要求，才能正常加载和运行 Outlook 加载项。

## <a name="client-requirements"></a>客户端要求

- 客户端必须是一个受 Outlook 加载项支持的主机。下列客户端支持加载项：

   - Windows 版 Outlook 2013 或更高版本
   - Mac 版 Outlook 2016 或更高版本
   - iOS 版 Outlook
   - Android 版 Outlook
   - 适用于 Exchange 2016 或更高版本和 Office 365 的 Outlook 网页版
   - 适用于 Exchange 2013 的 Outlook 网页版
   - Outlook.com

- The client must be connected to an Exchange server or Microsoft 365 using a direct connection. When configuring the client, the user must choose an **Exchange**, **Office 365**, or **Outlook.com** account type. If the client is configured to connect with POP3 or IMAP, add-ins will not load.

## <a name="mail-server-requirements"></a>邮件服务器要求

If the user is connected to Microsoft 365 or Outlook.com, mail server requirements are all taken care of already. However, for users connected to on-premises installations of Exchange Server, the following requirements apply.

- 服务器必须是 Exchange 2013 或更高版本。
- 必须启用 Exchange Web 服务 (EWS)，并向 Internet 公开此服务。 许多加载项要求，必须启用 EWS 才能正常运行。
- 服务器必须有有效身份验证证书，才能颁发有效标识令牌。 新安装的 Exchange Server 包含默认身份验证证书。 有关详细信息，请参阅 [Exchange 2016 中的数字证书和加密](/Exchange/architecture/client-access/certificates)和 [Set-AuthConfig](/powershell/module/exchange/organization/Set-AuthConfig)。
- 客户端访问服务器必须能够与 AppSource 通信，才能从 [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=a35323d5-0e3d-4cc0-ba44-57537d74aae8&omexanonuid=581941df-1c6f-4eda-89e7-651af8aeaeb2) 获取加载项。

## <a name="add-in-server-requirements"></a>加载项服务器要求

Add-in files (HTML, JavaScript, etc.) can be hosted on any web server platform desired. The only requirement is that the server must be configured to use HTTPS, and the SSL certificate must be trusted by the client.

## <a name="see-also"></a>另请参阅

- [Office 加载项的运行要求](../concepts/requirements-for-running-office-add-ins.md)
- [Office 加载项主机和平台可用性（Outlook 部分）](../overview/office-add-in-availability.md#outlook)
- [Outlook JavaScript API 要求集支持](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)
