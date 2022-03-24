---
title: 注册Office SSO 的加载项Microsoft 标识平台
description: 了解如何使用 Office 注册 Microsoft 标识平台 以将 SSO 与 Word、Excel、PowerPoint 和 Outlook 一Outlook。
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: e408a57534437f0d0fe0c5fb3b4ab844f7dde9ac
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/23/2022
ms.locfileid: "63743381"
---
# <a name="register-an-office-add-in-that-uses-single-sign-on-sso-with-the-microsoft-identity-platform"></a>在Office SSO 加载项中注册使用单一 (登录) 加载项Microsoft 标识平台

本文介绍如何在加载项Office加载项Microsoft 标识平台以便可以使用 SSO。 开始开发外接程序时注册它，以便当您继续测试或生产时，可以更改现有注册或为外接程序的开发、测试和生产版本创建单独的注册。

下表列出了执行此过程所需的信息以及说明中显示的相应占位符。

|信息  |示例  |占位符  |
|---------|---------|---------|
|加载项的人类可读名称。 （建议使用唯一名称，但不是必需的。）|`Contoso Marketing Excel Add-in (Prod)`|不适用|
|Azure 在注册过程中生成的应用程序 ID。|`c6c1f32b-5e55-4997-881a-753cc1d563b7`|`<application-id>`|
|加载项的完全限定域名（协议除外）。 *必须使用自己的域*。 正因如此，不能使用某些知名域名，例如 `azurewebsites.net` 或 `cloudapp.net`。 域必须相同，包括任何子域，如加载项清单的 `<Resources>` 部分中的 URL 中所使用的那样。|`localhost:6789`, `addins.contoso.com`|`<fully-qualified-domain-name>`|
|外接程序所需的 Microsoft 标识平台 和 Microsoft Graph权限。 （`profile` 始终是必需的。）|`profile`, `Files.Read.All`|不适用|

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]
