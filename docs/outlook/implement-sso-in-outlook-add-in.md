---
title: 应用场景 - 为服务实施单一登录
description: 了解如何使用 Outlook 加载项提供的单一登录令牌和 Exchange 标识令牌为服务实现 SSO。
ms.date: 09/03/2021
ms.localizationpriority: medium
ms.openlocfilehash: 2b9c4031a0011d2333582b4a10abe42f6844f763
ms.sourcegitcommit: 287a58de82a09deeef794c2aa4f32280efbbe54a
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 03/28/2022
ms.locfileid: "64496920"
---
# <a name="scenario-implement-single-sign-on-to-your-service-in-an-outlook-add-in"></a>应用场景：为 Outlook 加载项中的服务实现单一登录

在本文中，我们将探讨结合使用[单一登录访问令牌](authenticate-a-user-with-an-sso-token.md)和 [Exchange 标识令牌](authenticate-a-user-with-an-identity-token.md)为自己的后端服务提供单一登录实现的建议方法。 通过结合使用这两种令牌，可以在 SSO 访问令牌可用时利用其优势，并在其不可用时确保加载项仍能正常工作（例如，当用户切换到不支持这些令牌的客户端时，或当用户的邮箱位于本地 Exchange 服务器时）。

有关实现本文中的想法的示例外接程序，请参阅Outlook[外接程序 SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO)。


> [!NOTE]
> 目前，Word、Excel、Outlook 和 PowerPoint 支持单一登录 API。 若要详细了解目前支持单一登录 API 的平台，请参阅 [IdentityAPI 要求集](/javascript/api/requirement-sets/common/identity-api-requirement-sets)。 如果使用的是 Outlook 加载项，请务必为 Microsoft 365 租赁启用新式验证。 若要了解如何这样做，请参阅 [Exchange Online: 如何为租户启用新式验证](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)。


## <a name="why-use-the-sso-access-token"></a>为什么使用 SSO 访问令牌？

Exchange 标识令牌适用于加载项 API 的所有要求集，因此，仅依赖此令牌并完全忽略 SSO 令牌似乎是更好的做法。 但是，与 Exchange 标识令牌相比，SSO 令牌具有某些优势，因此，当该令牌可用时会建议使用此方法。

- SSO 令牌使用标准 OpenID 格式并由 Azure 颁发。 这极大地简化了验证这些令牌的过程。 与之相比，Exchange 标识令牌使用基于 JSON Web 令牌标准的自定义格式，因此需要使用自定义操作来验证此令牌。
- 后端可以使用 SSO 令牌来检索 Microsoft Graph 访问令牌，而用户无需执行任何其他登录操作。
- SSO 令牌提供的标识信息更为丰富，例如用户的显示名称。

## <a name="add-in-scenario"></a>加载项应用场景

鉴于此示例的目的，请考虑使用包含加载项 UI 和脚本 (HTML + JavaScript) 以及加载项调用的后端 Web API 的加载项。 后端 Web API 将同时调用 [Microsoft Graph API](/graph/overview) 和 Contoso 数据 API（由第三方创建的虚拟 API）。 与 Microsoft Graph API 类似，Contoso 数据 API 也需要进行 OAuth 身份验证。 要求是，后端 Web API 应能够同时调用这两个 API，而无需在每次访问令牌过期时提示用户提供凭据。

为了实现这一目的，后端 API 创建了一个安全的用户数据库。 每个用户都将在该数据库中获得一个条目，后端可以在其中存储 Microsoft Graph API 和 Contoso 数据 API 的长期刷新令牌。 以下 JSON 标记表示用户在数据库中的条目。

```JSON
{
  "userDisplayName": "...",
  "ssoId": "...",
  "exchangeId": "...",
  "graphRefreshToken": "...",
  "contosoRefreshToken": "..."
}
```

加载项会在对后端 Web API 的每个调用中包含 SSO 访问令牌（如果可用）或 Exchange 标识令牌（如果 SSO 令牌不可用）。

### <a name="add-in-startup"></a>加载项启动

1. 当加载项启动时，它向后端 Web API 发送请求，以确定用户是否已注册（即在用户数据库中是否有相关联的记录）以及 API 是否同时具有 Graph 和 Contoso 的刷新令牌。 在此调用中，加载项同时包含 SSO 令牌（如果可用）和标识令牌。

1. Web API 使用[使用 Outlook 加载项中的单一登录令牌对用户进行身份验证](authenticate-a-user-with-an-sso-token.md)和[使用 Exchange 标识令牌对用户进行身份验证](authenticate-a-user-with-an-identity-token.md)中的方法进行验证并从这两种令牌中生成唯一标识符。

1. 如果提供了 SSO 令牌，则 Web API 会查询用户数据库中是否存在具有 `ssoId` 值（该值与从 SSO 令牌生成的唯一标识符相匹配）的条目。
   - 如果条目不存在，则继续执行下一步。
   - 如果条目存在，则继续执行步骤 5。

1. Web API 查询数据库中是否存在具有 `exchangeId` 值（该值与从 Exchange 标识令牌生成的唯一标识符相匹配）的条目。
   - 如果该条目存在且 SSO 令牌可用，则更新该数据库中的用户记录，以将 `ssoId` 值设置为从 SSO 令牌生成的唯一标识符，并继续执行步骤 5。
   - 如果该条目存在但 SSO 令牌不可用，则继续执行步骤 5。
   - 如果该条目不存在，则新建一个条目。 将 `ssoId` 设置为从 SSO 令牌生成的唯一标识符（如果可用），并将 `exchangeId` 设置为从 Exchange 标识令牌生成的唯一标识符。

1. 检查用户的 `graphRefreshToken` 值中是否存在有效的刷新令牌。
   - 如果此值缺失或无效且 SSO 令牌可用，则使用 [OAuth2代表流](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)获取 Graph 的访问令牌和刷新令牌。 将刷新令牌保存在用户的 `graphRefreshToken` 值中。

1. 检查 `graphRefreshToken` 和 `contosoRefreshToken` 中是否存在有效的刷新令牌。
   - 如果两个值均有效，则对加载项做出响应，来指示用户已注册且已进行了配置。
   - 如果任一值无效，则对加载项做出响应，来指示需要进行用户设置，并指示需要配置的服务（Graph 和 Contoso）。

1. 加载项将检查响应。
   - 如果用户已注册并已进行了配置，则加载项将继续正常运行。
   - 如需进行用户设置，则加载项进入“设置”模式并提示用户向加载项授权。

### <a name="authorize-the-backend-web-api"></a>授权后端 Web API

理想情况下，授权后端 Web API 调用 Microsoft Graph API 和 Contoso 数据 API 这一过程应仅进行一次，以尽量减少提示用户进行登录的次数。

基于后端 Web API 的响应，加载项可能需要授权用户使用 Microsoft Graph API 和/或 Contoso 数据 API。 因为这两种 API 都使用 OAuth2 身份验证，所以它们的授权方法类似。

1. 加载项通知用户需要授权其使用 API 并让用户单击一个链接或按钮来启动这一过程。

    > [!NOTE]
    > [Outlook 外接程序 SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO) 的示例外接程序演示如何使用[对话框 API](/javascript/api/office/office.ui#displaydialogasync-startaddress--options--callback-) 和 [office-js-helpers](https://github.com/OfficeDev/office-js-helpers) 库作为选项来启动 [API 的 OAuth2 授权](/azure/active-directory/develop/active-directory-protocols-oauth-code)代码流。

1. 此流完成后，加载项向后端 Web API 发送刷新令牌并包含 SSO 令牌（如果可用）或 Exchange 标识令牌。

1. 后端 Web API 在数据库中查找用户并更新相应的刷新令牌。

1. 加载项将继续正常运行。

### <a name="normal-operation"></a>正常运行

每当加载项调用后端 Web API 时，它都将包含 SSO 令牌或 Exchange 标识令牌。 后端 Web API 根据此令牌查找用户，然后使用存储的刷新令牌来获取 Microsoft Graph API 和 Contoso 数据 API 的访问令牌。 只要刷新令牌有效，用户就无需再次登录。
