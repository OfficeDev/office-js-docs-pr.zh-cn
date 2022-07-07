---
title: 将使用 SSO 的 Office 加载项注册到Microsoft 标识平台
description: 了解如何向Microsoft 标识平台注册 Office 加载项，以便将 SSO 与 Word、Excel、PowerPoint 和 Outlook 配合使用。
ms.date: 01/25/2022
ms.localizationpriority: medium
ms.openlocfilehash: 69506c4b98da2e7d70e82cf49093a75374e77f92
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659778"
---
# <a name="register-an-office-add-in-that-uses-single-sign-on-sso-with-the-microsoft-identity-platform"></a>将使用单一登录 (SSO) 的 Office 加载项注册到Microsoft 标识平台

本文介绍如何向Microsoft 标识平台注册 Office 加载项，以便可以使用 SSO。 在开始开发外接程序时注册外接程序，以便在进行测试或生产时，可以更改现有注册或创建单独的注册，以便开发、测试和生产版本的外接程序。

下表列出了执行此过程所需的信息以及说明中显示的相应占位符。

|信息  |示例  |占位符  |
|---------|---------|---------|
|加载项的人类可读名称。 （建议使用唯一名称，但不是必需的。）|`Contoso Marketing Excel Add-in (Prod)`|不适用|
|Azure 在注册过程中为你生成的应用程序 ID。|`c6c1f32b-5e55-4997-881a-753cc1d563b7`|`<application-id>`|
|加载项的完全限定域名（协议除外）。 *必须使用自己的域*。 正因如此，不能使用某些知名域名，例如 `azurewebsites.net` 或 `cloudapp.net`。 域必须相同，包括任何子域，如外接程序清单部分的 URL 中 **\<Resources\>** 所用。|`localhost:6789`, `addins.contoso.com`|`<fully-qualified-domain-name>`|
|加载项所需的对 Microsoft 标识平台 和 Microsoft Graph 的权限。 （`profile` 始终是必需的。）|`profile`, `Files.Read.All`|不适用|

[!INCLUDE[](../includes/register-sso-add-in-aad-v2-include.md)]
