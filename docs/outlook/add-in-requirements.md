---
title: Outlook 加载项要求
description: 必须满足服务器和客户端的多个要求，才能正常加载和运行 Outlook 加载项。
ms.date: 02/09/2021
localization_priority: Priority
ms.openlocfilehash: 6062073d44a412d67961f806677cd60701bbdb9b
ms.sourcegitcommit: 42c55a8d8e0447258393979a09f1ddb44c6be884
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/08/2021
ms.locfileid: "58936552"
---
# <a name="outlook-add-in-requirements"></a>Outlook 加载项要求

必须满足服务器和客户端的多个要求，才能正常加载和运行 Outlook 加载项。

## <a name="client-requirements"></a>客户端要求

- 客户端必须是一个受支持的 Outlook 加载项应用程序。下列客户端支持加载项。

  - Windows 版 Outlook 2013 或更高版本
  - Mac 版 Outlook 2016 或更高版本
  - iOS 版 Outlook
  - Android 版 Outlook
  - 适用于 Exchange 2016 或更高版本的 Outlook 网页版
  - 适用于 Exchange 2013 的 Outlook 网页版
  - Outlook.com

- 必须使用直接连接将客户端连接到 Exchange 服务器或 Microsoft 365。配置客户端时，用户必须选择 **Exchange**、**Office** 或 **Outlook.com** 帐户类型。如果将客户端配置为使用 POP3 或 IMAP 连接，则加载项将不会加载。

## <a name="mail-server-requirements"></a>邮件服务器要求

如果用户已连接到 Microsoft 365 或 Outlook.com，则已经满足了所有邮件服务器要求。但是，对于连接到 Exchange Server 本地安装的用户，适用以下要求。

- 服务器必须是 Exchange 2013 或更高版本。
- 必须启用 Exchange Web 服务 (EWS)，并向 Internet 公开此服务。 许多加载项要求，必须启用 EWS 才能正常运行。
- 服务器必须有有效身份验证证书，才能颁发有效标识令牌。 新安装的 Exchange Server 包含默认身份验证证书。 有关详细信息，请参阅 [Exchange 2016 中的数字证书和加密](/Exchange/architecture/client-access/certificates)和 [Set-AuthConfig](/powershell/module/exchange/organization/Set-AuthConfig)。
- 客户端访问服务器必须能够与 AppSource 通信，才能从 [AppSource](https://appsource.microsoft.com/marketplace/apps?product=office&page=1&src=office&corrid=a35323d5-0e3d-4cc0-ba44-57537d74aae8&omexanonuid=581941df-1c6f-4eda-89e7-651af8aeaeb2) 获取加载项。

## <a name="add-in-server-requirements"></a>加载项服务器要求

可在任意需要的 Web 服务器平台上托管外接程序文件（HTML、JavaScript 等）。唯一的要求是，必须将服务器配置为使用 HTTPS，并且 SSL 证书必须受客户端信任。

## <a name="see-also"></a>另请参阅

- [Office 加载项的运行要求](../concepts/requirements-for-running-office-add-ins.md)
- [Office 客户端应用程序和 Office 加载项的平台可用性（Outlook 部分）](../overview/office-add-in-availability.md#outlook)
- [Outlook JavaScript API 要求集支持](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients)
