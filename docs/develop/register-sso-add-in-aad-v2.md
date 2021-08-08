---
title: 向 Azure AD v2.0 端点注册使用 SSO 的 Office 加载项。
description: 了解如何使用 Azure AD v2.0 Office注册加载项。
ms.date: 04/10/2019
localization_priority: Normal
ms.openlocfilehash: c7ae397fba0bf92e5ef2ef8795ef12cd2036ed65fac5e18c19521b342998f03f
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 08/07/2021
ms.locfileid: "57080193"
---
# <a name="register-an-office-add-in-that-uses-sso-with-the-azure-ad-v20-endpoint"></a>向 Azure AD v2.0 端点注册使用 SSO 的 Office 加载项。

本文介绍如何向 Azure AD v2.0 端点注册 Office 加载项。 开始开发时，需要注册加载项。 在进行测试或生产时，可以更改现有注册或为加载项的开发、测试和生产版本创建单独的注册。

下表列出了执行此过程所需的信息以及说明中显示的相应占位符。

|信息  |示例  |占位符  |
|---------|---------|---------|
|加载项的人类可读名称。 （建议使用唯一名称，但不是必需的。）|`Contoso Marketing Excel Add-in (Prod)`|**$ADD-IN-NAME$**|
|加载项的完全限定域名（协议除外）。 *必须使用自己的域*。 正因如此，不能使用某些知名域名，例如 `azurewebsites.net` 或 `cloudapp.net`。 域必须相同，包括任何子域，如加载项清单的 `<Resources>` 部分中的 URL 中所使用的那样。|`localhost:6789`, `addins.contoso.com`|**$FQDN-WITHOUT-PROTOCOL$**|
|加载项所需的 AAD 和 Microsoft Graph 权限。 （`profile` 始终是必需的。）|`profile`, `Files.Read.All`|不适用|

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]
