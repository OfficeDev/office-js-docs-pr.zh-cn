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
# <a name="register-an-office-add-in-that-uses-sso-with-the-azure-ad-v20-endpoint"></a>向 Azure AD v2.0 端点注册使用 SSO 的 Office 外接程序

本文介绍了如何向 Azure AD v2.0 端点注册 Office 外接程序。 开始开发时，需要注册外接程序。 进行测试或生产时，可以为外接 程序的开发、测试和生产版本更改现有注册或创建单独的注册。 

下表列出了执行此过程所需的信息以及说明中出现的相应占位符。 

|信息  |示例  |占位符  |
|---------|---------|---------|
|外接程序的人类可读名称。 （建议使用唯一名称，但并非强制性要求。）    |`Contoso Marketing Excel Add-in (Prod)`        |**$ADD-IN-NAME$**         |
|外接 程序的完全限定的域名（协议除外）。 *必须使用你所拥有的域名。* 出于这个原因，你不能使用某些众所周知的领域，如 `azurewebsites.net` 或者 `cloudapp.net`。   |`localhost:6789`, `addins.contoso.com`         |**$FQDN-WITHOUT-PROTOCOL$**         |
|外接程序所需的 AAD 和 Microsoft Graph 权限。 （始终需要 `profile`。）    |`profile`, `Files.Read.All`         |无         |

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]