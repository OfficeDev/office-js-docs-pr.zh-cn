---
title: 使用 SSO 对 Microsoft Graph 授权
description: 了解 Office 外接程序的用户可以如何使用单一登录 (SSO) 从 Microsoft Graph 获取数据。
ms.date: 07/30/2020
localization_priority: Normal
ms.openlocfilehash: e87c86b5302bde8122485b837759fa327251c656
ms.sourcegitcommit: ceb8dd66f3fb9c963fce8446c2f6c65ead56fbc1
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 11/18/2020
ms.locfileid: "49131911"
---
# <a name="authorize-to-microsoft-graph-with-sso"></a>使用 SSO 对 Microsoft Graph 授权

用户可以使用自己的个人 Microsoft 帐户或 Microsoft 365 教育或工作帐户，登录 Office（在线、移动和桌面平台）。 在 Office 加载项中获取对 [Microsoft Graph](https://developer.microsoft.com/graph/docs) 的访问权限的最佳方式是使用用户的 Office 登录凭据。 这使用户能够访问其 Microsoft Graph 数据，而无需再次登录。

> [!NOTE]
> 目前，Word、Excel、Outlook 和 PowerPoint 支持单一登录 API。 若要详细了解目前支持单一登录 API 的平台，请参阅 [IdentityAPI 要求集](/office/dev/add-ins/reference/requirement-sets/identity-api-requirement-sets)。
> 如果使用的是 Outlook 加载项，请务必为 Office 365 租赁启用新式验证。 若要了解如何这样做，请参阅 [Exchange Online：如何为租户启用新式验证](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)。

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>SSO 和 Microsoft Graph 的加载项体系结构

除了托管 Web 应用程序的页面和 JavaScript 之外，外接程序还必须以相同的[完全限定的域名](/windows/desktop/DNS/f-gly#_dns_fully_qualified_domain_name_fqdn__gly)托管一个或多个 Web API，这些 API 可获取 Microsoft Graph 的访问令牌，并向它发出请求。

外接程序清单包含标记，用于指定外接程序在 Azure Active Directory (Azure AD) v2.0 终结点中的注册方式，并指定外接程序需要的 Microsoft Graph 的任何权限。

### <a name="how-it-works-at-runtime"></a>运行时的工作方式

下图展示了 Microsoft Graph 登录和访问流程的工作原理。

![显示 SSO 过程的关系图](../images/sso-access-to-microsoft-graph.png)

1. 在加载项中，JavaScript 调用新的 Office.js API [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-)。 该操作告诉 Office 客户端应用程序获取加载项的访问令牌。 （以下称为 **启动访问令牌**，因为在该过程的后期它将会被替换为另一个令牌。 有关已解码启动访问令牌的示例，请参阅[示例访问令牌](sso-in-office-add-ins.md#example-access-token)。）
2. 如果用户未登录，Office 客户端应用程序会打开弹出窗口，以供用户登录。
3. 如果当前用户是首次使用加载项，则会看到同意提示。
4. Office 客户端应用程序从当前用户的 Azure AD v2.0 终结点请求 **启动访问令牌** 。
5. Azure AD 将引导令牌发送到 Office 客户端应用程序。
6. Office 客户端应用程序将 **引导访问令牌** 作为调用返回的 result 对象的一部分发送到外接程序 `getAccessToken` 。
7. 加载项中的 JavaScript 向 Web API（与加载项托管在同一完全限定的域中）发出 HTTP 请求，并添加 **启动访问令牌** 作为授权证明。
8. 服务器端代码验证传入的 **启动访问令牌**。
9. 服务器端代码使用 "代表" 流 (在 [OAuth2 令牌交换](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02) 和 [守护程序或服务器应用程序上定义到 web API Azure 方案](/azure/active-directory/develop/active-directory-authentication-scenarios)) ，以获取 Exchange 中的 Microsoft Graph 访问令牌，以获取对启动访问令牌的 Exchange。
10. Azure AD 将 Microsoft Graph 访问令牌（如果加载项请求获取 *offline_access* 权限，则同时返回刷新令牌）返回给加载项。
11. 服务器端代码缓存 Microsoft Graph 访问令牌。
12. 服务器端代码向 Microsoft Graph 发出请求，并添加 Microsoft Graph 访问令牌。
13. Microsoft Graph 将数据返回到加载项，该加载项可将其传递到加载项的 UI。
14. 当 Microsoft Graph 访问令牌过期时，服务器端代码可以使用其刷新令牌获取新的 Microsoft Graph 访问令牌。

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>开发可访问 Microsoft Graph 的 SSO 加载项

开发一个可访问 Microsoft Graph 的加载项，就像可使用 SSO 的任何其他加载项一样。 有关完整的说明，请参阅[为 Office 加载项启用单一登录](../develop/sso-in-office-add-ins.md)。区别在于，加载项必须具有服务器端 Web API，并且我们将该文中的访问令牌成为“启动访问令牌”。

根据所用的语言和框架，可能存在一些库，可简化必须编写的服务器端代码。 代码应执行以下操作：

* 使用对 Azure AD v2.0 终结点的调用（包括 "启动" 访问令牌、有关用户的一些元数据和外接程序的凭据 (ID 和机密) ）启动 "代表" 流。
* 创建一个或多个 Web API 方法，用于通过将可能缓存的访问令牌传递给 Microsoft Graph 来获取 Microsoft Graph 数据。
* 或者，在启动流程之前，验证从之前创建的令牌处理程序收到的加载项启动访问令牌。 有关详细信息，请参阅[验证访问令牌](sso-in-office-add-ins.md#validate-the-access-token)。 
* 或者，在流程完成后，将返回的访问令牌缓存到 Microsoft Graph。 如果加载项对 Microsoft Graph 进行多次调用，则需要执行此操作。 有关此流的详细信息，请参阅 [Azure Active Directory v2.0 和 OAuth 2.0 代表流](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)。

> [!NOTE]
> 有关“代表”流获取的 Microsoft Graph 已解码访问令牌的示例，请参阅 [Azure Active Directory v2.0 和 OAuth 2.0 代表流](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)。

有关详细演练和应用场景的示例，请参阅：

* [创建使用单一登录的 Node.js Office 加载项](create-sso-office-add-ins-nodejs.md)
* [创建使用单一登录的 ASP.NET Office 加载项](create-sso-office-add-ins-aspnet.md)
* [应用场景：为 Outlook 加载项中的服务实现单一登录](../outlook/implement-sso-in-outlook-add-in.md)

## <a name="distributing-sso-enabled-add-ins-in-microsoft-appsource"></a>在 Microsoft AppSource 中分发启用了 SSO 的外接程序

当 Microsoft 365 管理员从 [AppSource](https://appsource.microsoft.com)获取加载项时，管理员可以通过 [集中部署](../publish/centralized-deployment.md) 来重新发布它，并向外接程序授予管理员同意，以访问 Microsoft Graph 作用域。 但是，最终用户也可以直接从 AppSource 获取外接程序，在这种情况下，用户必须向外接程序授予许可。 这可能会带来潜在的性能问题，我们为其提供了解决方案。

如果您 `allowConsentPrompt` 的代码在的调用中传递选项 `getAccessToken` （例如 `OfficeRuntime.auth.getAccessToken( { allowConsentPrompt: true } );` ），则 Office 可以在 Azure AD 向 office 报告同意尚未授予外接程序的情况下提示用户同意。 但是，出于安全考虑，Office 只会提示用户同意 Azure AD `profile` 作用域。 *Office 不会提示同意任何 Microsoft Graph 作用域*，甚至不是偶数 `User.Read` 。 这意味着，如果用户向提示授予许可，Office 将返回一个引导令牌。 但是，尝试将访问令牌的引导令牌交换到 Microsoft Graph 的尝试将会失败，并出现错误 AADSTS65001，这意味着尚未授予对 Microsoft Graph 作用域) 的许可 (。

您的代码可以，并应通过回退到备用的身份验证系统来处理此错误，这将提示用户同意 Microsoft Graph 作用域。  (有关代码示例的详细说明，请参阅 [创建使用单一登录的 Node.js Office 加载项](create-sso-office-add-ins-nodejs.md) ，并 [创建使用单一登录的 ASP.NET office 加载](create-sso-office-add-ins-aspnet.md) 项以及它们所链接到的示例。 ) 但整个过程需要多个到 Azure AD 的往返行程。 您可以通过 `forMSGraphAccess` 在调用中包括选项来避免这种性能下降 `getAccessToken` ，例如， `OfficeRuntime.auth.getAccessToken( { forMSGraphAccess: true } )` 。  这会向 Office 发出通知，指示你的外接程序需要 Microsoft Graph 作用域。 Office 将要求 Azure AD 验证是否已向外接程序授予 Microsoft Graph 作用域的同意。 如果已有，则将返回引导令牌。 如果没有，则调用 `getAccessToken` 将返回错误13012。 您的代码可以立即回退到备用的身份验证系统来处理此错误，而无需进行 doomed 尝试与 Azure AD 交换令牌。

最佳做法是，始终传递 `forMSGraphAccess` 给 `getAccessToken` 您的外接程序将在 AppSource 中分发，并需要 Microsoft Graph 作用域。

> [!TIP]
> 如果开发使用 SSO 的 Outlook 外接程序并旁加载它进行测试，则即使已 *always* `forMSGraphAccess` `getAccessToken` 授予管理员同意，Office 也始终会返回错误13012。 因此，在 `forMSGraphAccess` 开发 Outlook 外接程序 **时** ，应注释掉该选项。 在部署生产时，请务必取消对该选项的注释。 仅当您在 Outlook 中进行旁加载时，才会发生虚假13012。
