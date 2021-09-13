---
title: Outlook 加载项中的身份验证选项
description: Outlook 加载项 根据特定场景提供了多种不同的身份验证方法。
ms.date: 09/03/2021
ms.localizationpriority: high
ms.openlocfilehash: c51f40cda48fbb343cfae78b258902e5030f49da
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: HT
ms.contentlocale: zh-CN
ms.lasthandoff: 09/12/2021
ms.locfileid: "59152431"
---
# <a name="authentication-options-in-outlook-add-ins"></a>Outlook 加载项中的身份验证选项

Outlook 加载项可以访问 Internet 上任意位置的信息，无论是托管加载项的服务器、内部网络，还是云中的其他位置。 如果相应信息受保护，加载项需要能够验证用户身份。 Outlook 加载项 根据特定场景提供了多种不同的身份验证方法。

## <a name="single-sign-on-access-token"></a>单一登录访问令牌

单一登录访问令牌为你的加载项提供了进行身份验证和获取访问令牌以调用 [Microsoft Graph API](/graph/overview) 的无缝方法。 由于不需要用户输入其凭据，此功能可以减少摩擦。

> [!NOTE]
> 目前，Word、Excel、Outlook 和 PowerPoint 支持单一登录 API。 若要详细了解目前支持单一登录 API 的平台，请参阅 [IdentityAPI 要求集](../reference/requirement-sets/identity-api-requirement-sets.md)。
> 如果使用的是 Outlook 加载项，请务必为 Microsoft 365 租赁启用新式验证。 若要了解如何执行此操作，请参阅 [Exchange Online: How to enable your tenant for modern authentication](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)（如何为租户启用新式体验）。

如果加载项符合以下情况，请考虑使用 SSO 访问令牌：

- 主要由 Microsoft 365 用户使用
- 需要访问以下服务：
  - 作为 Microsoft Graph 的一部分公开的 Microsoft 服务
  - 你控制的非 Microsoft 服务

SSO 身份验证方法使用 [Azure Active Directory 提供的 OAuth2 代表流](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)。 它要求加载项在[应用程序注册门户](https://apps.dev.microsoft.com/)中进行注册并在其清单中指定任何所需的 Microsoft Graph 作用域。

借助此方法，加载项可以获取作用域为你的服务器后端 API 的访问令牌。 加载项将此令牌用作 `Authorization` 标头中的持有者令牌，来对 API 回调进行身份验证。 此时，服务器可以：

- 完成“代表”流来获取作用域为 Microsoft Graph API 的访问令牌
- 使用令牌中的标识信息创建用户标识并对自己的后端服务进行身份验证

有关更详细的概述，请参阅 [SSO 身份验证方法的完整概述](../develop/sso-in-office-add-ins.md)。

有关在 Outlook 加载项中使用 SSO 令牌的详细信息，请参阅[在 Outlook 加载项中使用单一登录令牌对用户进行身份验证](authenticate-a-user-with-an-sso-token.md)。

有关使用 SSO 令牌的加载项示例，请参阅 [Outlook 加载项 SSO](https://github.com/OfficeDev/PnP-OfficeAddins/tree/main/Samples/auth/Outlook-Add-in-SSO)。

## <a name="exchange-user-identity-token"></a>Exchange 用户标识令牌

Exchange 用户标识令牌为加载项提供了一种创建用户标识的方法。 通过验证用户标识，可以对后端系统执行一次性身份验证，然后接受用户标识令牌，来作为对未来请求的授权。 使用 Exchange 用户标识令牌：

- 当加载项主要由 Exchange 本地用户使用时。
- 当加载项需要访问你控制的非 Microsoft 服务时。
- 当加载项在不支持 SSO 的 Office 版本上运行时，要回退身份验证。

加载项可以调用 [getUserIdentityTokenAsync](/javascript/api/outlook/office.mailbox#getCallbackTokenAsync_callback__userContext_) 以获取 Exchange 用户标识令牌。 有关使用这些令牌的详细信息，请参阅[使用 Exchange 标识令牌对用户进行身份验证](authenticate-a-user-with-an-identity-token.md)。

## <a name="access-tokens-obtained-via-oauth2-flows"></a>通过 OAuth2 流获取的访问令牌

加载项也可以访问支持 OAuth2 进行授权的第三方服务。 如果你的加载项符合以下情况，请考虑使用 OAuth2 令牌：

- 需要访问不受你控制的第三方服务

使用此方法，加载项会提示用户通过使用 [displayDialogAsync](/javascript/api/office/office.ui#displayDialogAsync_startAddress__options__callback_) 方法初始化 OAuth2 流或使用 [office-js-helpers 库](https://github.com/OfficeDev/office-js-helpers) 转到 OAuth2 隐式流来登录到服务。

## <a name="callback-tokens"></a>回调令牌

借助回调令牌，可以使用 [Exchange Web 服务 (EWS)](/exchange/client-developer/exchange-web-services/explore-the-ews-managed-api-ews-and-web-services-in-exchange) 或 [Outlook REST API](/previous-versions/office/office-365-api/api/version-2.0/use-outlook-rest-api) 从服务器后端访问用户邮箱。 如果你的加载项符合以下情况，请考虑使用回调令牌：

- 需要从服务器后端访问用户邮箱。

加载项使用 [getCallbackTokenAsync ](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods)方法之一获取回调令牌。 访问权限级别由加载项清单中指定的权限控制。
