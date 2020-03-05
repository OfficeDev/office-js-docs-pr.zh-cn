---
title: 使用单一登录令牌对用户进行身份验证
description: 了解如何使用 Outlook 外接程序提供的单一登录令牌为服务实现 SSO。
ms.date: 11/19/2019
localization_priority: Normal
ms.openlocfilehash: 13babbb417618005f55e965d54e5b497d35bf6af
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166034"
---
# <a name="authenticate-a-user-with-a-single-sign-on-token-in-an-outlook-add-in-preview"></a>在 Outlook 加载项中使用单一登录令牌验证用户（预览版）

使用单一登录 (SSO)，加载项可以无缝方式验证用户（并根据需要获取访问令牌来调用 [Microsoft Graph API](/graph/overview)）。

借助此方法，加载项可以获取范围限定为服务器后端 API 的访问令牌。 加载项将此令牌用作 `Authorization` 头中的持有者令牌，以验证 API 回调。 也可以使用服务器端代码执行以下操作：

- 完成“代表”流，以获取范围限定为 Microsoft Graph API 的访问令牌
- 使用令牌中的标识信息，以创建用户标识并验证自己的后端服务

有关 Office 外接程序中的 SSO 的概述，请参阅[为 Office 外接程序启用单一登录](../develop/sso-in-office-add-ins.md)和[在 Office 外接程序中授予对 Microsoft Graph 的访问权限](../develop/authorize-to-microsoft-graph.md)。

> [!NOTE]
> 若要使用 SSO，必须从外接程序的启动 HTML 页面中的 https://appsforoffice.microsoft.com/lib/beta/hosted/office.js 加载 Office JavaScript 库的 Beta 版。

## <a name="enable-modern-authentication-in-your-office-365-tenancy"></a>在 Office 365 租赁中启用新式验证

必须为 Office 365 租赁启用新式验证，才能将 SSO 与 Outlook 外接程序结合使用。 若要了解如何执行此操作，请参阅 [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)（如何为租户启用新式体验）。

## <a name="register-your-add-in"></a>注册外接程序

若要使用 SSO，Outlook 外接程序需要有已向 Azure Active Directory (AAD) v2.0 注册的服务器端 Web API。 有关详细信息，请参阅[向 Azure AD v2.0 端点注册使用 SSO 的 Office 外接程序](../develop/register-sso-add-in-aad-v2.md)。

### <a name="provide-consent-when-sideloading-an-add-in"></a>旁加载加载项时授予许可

从 AppSource 获取使用 SSO 的加载项时，应用商店 UI 将负责提示用户同意授予所请求的 Graph 权限。 但是，在开发加载项时，需要提前提供授权。 有关详细信息，请参阅[向加载项授予管理员许可](../develop/grant-admin-consent-to-an-add-in.md)

## <a name="update-the-add-in-manifest"></a>更新加载项清单

若要在加载项中启用 SSO，下一步在 `VersionOverridesV1_1` [VersionOverrides](../reference/manifest/versionoverrides.md) 元素末尾添加 `WebApplicationInfo` 元素。 有关详细信息，请参阅[配置加载项](../develop/sso-in-office-add-ins.md#configure-the-add-in)。

## <a name="get-the-sso-token"></a>获取 SSO 令牌

加载项使用客户端脚本获取 SSO 令牌。 有关详细信息，请参阅[添加客户端代码](../develop/sso-in-office-add-ins.md#add-client-side-code)。

## <a name="use-the-sso-token-at-the-back-end"></a>在后端使用 SSO 令牌

大多数情况下，如果加载项没有将访问令牌传递到服务器端并在其中使用它，那么获取访问令牌的意义就不大。 若要详细了解服务器端可以和应该执行的操作，请参阅[添加服务器端代码](../develop/sso-in-office-add-ins.md#add-server-side-code)。

> [!IMPORTANT]
> 若要将 SSO 令牌用作 *Outlook* 加载项中的标识，建议还[使用 Exchange 标识令牌](authenticate-a-user-with-an-identity-token.md)作为备用标识。 加载项用户可能使用多个客户端，而有些客户端可能不支持提供 SSO 令牌。 通过将 Exchange 标识令牌用作备用令牌，就不用多次提示这些用户输入凭据了。 有关详细信息，请参阅[应用场景：在 Outlook 外接程序中对服务实现单一登录](implement-sso-in-outlook-add-in.md)。

## <a name="see-also"></a>另请参阅

- 有关使用 SSO 令牌访问 Microsoft Graph API 的 Outlook 外接程序示例，请参阅 [AttachmentsDemo 示例外接程序](https://github.com/OfficeDev/outlook-add-in-attachments-demo)。
- [SSO API 参考](../develop/sso-in-office-add-ins.md#sso-api-reference)
- [IdentityAPI 要求集](../reference/requirement-sets/identity-api-requirement-sets.md)