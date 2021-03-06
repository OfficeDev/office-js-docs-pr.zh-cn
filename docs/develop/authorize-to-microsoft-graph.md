---
title: 使用 SSO 对 Microsoft Graph 授权
description: 了解 Office 外接程序的用户如何使用单一登录 (SSO) 从 Microsoft Graph 获取数据。
ms.date: 02/09/2021
localization_priority: Normal
ms.openlocfilehash: 2f72b19023d9c5fdb8e35466bbd64269cbab81ec
ms.sourcegitcommit: ccc0a86d099ab4f5ef3d482e4ae447c3f9b818a3
ms.translationtype: MT
ms.contentlocale: zh-CN
ms.lasthandoff: 02/14/2021
ms.locfileid: "50237861"
---
# <a name="authorize-to-microsoft-graph-with-sso"></a>使用 SSO 对 Microsoft Graph 授权

用户可以使用自己的个人 Microsoft 帐户或 Microsoft 365 教育或工作帐户，登录 Office（在线、移动和桌面平台）。 在 Office 加载项中获取对 [Microsoft Graph](https://developer.microsoft.com/graph/docs) 的访问权限的最佳方式是使用用户的 Office 登录凭据。 这使用户能够访问其 Microsoft Graph 数据，而无需再次登录。

> [!NOTE]
> 目前，Word、Excel、Outlook 和 PowerPoint 支持单一登录 API。 若要详细了解目前支持单一登录 API 的平台，请参阅 [IdentityAPI 要求集](../reference/requirement-sets/identity-api-requirement-sets.md)。
> 如果使用的是 Outlook 加载项，请务必为 Microsoft 365 租赁启用新式验证。 若要了解如何这样做，请参阅 [Exchange Online：如何为租户启用新式验证](https://social.technet.microsoft.com/wiki/contents/articles/32711.exchange-online-how-to-enable-your-tenant-for-modern-authentication.aspx)。

## <a name="add-in-architecture-for-sso-and-microsoft-graph"></a>SSO 和 Microsoft Graph 的加载项体系结构

除了托管 Web 应用程序的页面和 JavaScript 之外，外接程序还必须以相同的[完全限定的域名](/windows/desktop/DNS/f-gly#_dns_fully_qualified_domain_name_fqdn__gly)托管一个或多个 Web API，这些 API 可获取 Microsoft Graph 的访问令牌，并向它发出请求。

外接程序清单包含标记，用于指定外接程序在 Azure Active Directory (Azure AD) v2.0 终结点中的注册方式，并指定外接程序需要的 Microsoft Graph 的任何权限。

### <a name="how-it-works-at-runtime"></a>运行时的工作方式

下图展示了 Microsoft Graph 登录和访问流程的工作原理。

![显示 SSO 过程的图表](../images/sso-access-to-microsoft-graph.png)

1. 在加载项中，JavaScript 调用新的 Office.js API [getAccessToken](/javascript/api/office-runtime/officeruntime.auth#getaccesstoken-options-)。 该操作告诉 Office 客户端应用程序获取加载项的访问令牌。 （以下称为 **启动访问令牌**，因为在该过程的后期它将会被替换为另一个令牌。 有关已解码启动访问令牌的示例，请参阅[示例访问令牌](sso-in-office-add-ins.md#example-access-token)。）
2. 如果用户未登录，Office 客户端应用程序会打开弹出窗口，以供用户登录。
3. 如果当前用户是首次使用加载项，则会看到同意提示。
4. Office 客户端应用程序从 Azure  AD v2.0 终结点请求当前用户的启动访问令牌。
5. Azure AD 将启动令牌发送到 Office 客户端应用程序。
6. Office 客户端应用程序将启动访问令牌作为调用返回的结果对象的一 `getAccessToken` 部分发送到外接程序。
7. 加载项中的 JavaScript 向 Web API（与加载项托管在同一完全限定的域中）发出 HTTP 请求，并添加 **启动访问令牌** 作为授权证明。
8. 服务器端代码验证传入的 **启动访问令牌**。
9. 服务器端代码使用 [OAuth2](https://tools.ietf.org/html/draft-ietf-oauth-token-exchange-02) 令牌 Exchange 中定义的"代表"流 (以及 Web [API Azure](/azure/active-directory/develop/active-directory-authentication-scenarios) 方案的守护程序或服务器应用程序) 来获取 Microsoft Graph 的访问令牌，以交换启动访问令牌。
10. Azure AD 将 Microsoft Graph 访问令牌（如果加载项请求获取 *offline_access* 权限，则同时返回刷新令牌）返回给加载项。
11. 服务器端代码缓存 Microsoft Graph 访问令牌。
12. 服务器端代码向 Microsoft Graph 发出请求，并添加 Microsoft Graph 访问令牌。
13. Microsoft Graph 将数据返回到加载项，加载项可以将其传递到加载项的 UI。
14. 当 Microsoft Graph 访问令牌过期时，服务器端代码可以使用其刷新令牌获取新的 Microsoft Graph 访问令牌。

## <a name="develop-an-sso-add-in-that-accesses-microsoft-graph"></a>开发可访问 Microsoft Graph 的 SSO 加载项

开发一个可访问 Microsoft Graph 的加载项，就像可使用 SSO 的任何其他加载项一样。 有关完整的说明，请参阅[为 Office 加载项启用单一登录](../develop/sso-in-office-add-ins.md)。区别在于，加载项必须具有服务器端 Web API，并且我们将该文中的访问令牌成为“启动访问令牌”。

根据所用的语言和框架，可能存在一些库，可简化必须编写的服务器端代码。 代码应执行以下操作：

* 通过调用 Azure AD v2.0 终结点启动"代表"流，该终结点包括启动访问令牌、有关用户的一些元数据以及外接程序的凭据 (其 ID 和密码) 。
* 创建一个或多个 Web API 方法，用于通过将可能缓存的访问令牌传递给 Microsoft Graph 来获取 Microsoft Graph 数据。
* 或者，在启动流程之前，验证从之前创建的令牌处理程序收到的加载项启动访问令牌。 有关详细信息，请参阅[验证访问令牌](sso-in-office-add-ins.md#validate-the-access-token)。 
* 或者，在流程完成后，将返回的访问令牌缓存到 Microsoft Graph。 如果加载项对 Microsoft Graph 进行多次调用，则需要执行此操作。 有关此流的详细信息，请参阅 [Azure Active Directory v2.0 和 OAuth 2.0 代表流](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)。

> [!NOTE]
> 有关“代表”流获取的 Microsoft Graph 已解码访问令牌的示例，请参阅 [Azure Active Directory v2.0 和 OAuth 2.0 代表流](/azure/active-directory/develop/active-directory-v2-protocols-oauth-on-behalf-of)。

有关详细演练和应用场景的示例，请参阅：

* [创建使用单一登录的 Node.js Office 加载项](create-sso-office-add-ins-nodejs.md)
* [创建使用单一登录的 ASP.NET Office 加载项](create-sso-office-add-ins-aspnet.md)
* [应用场景：为 Outlook 加载项中的服务实现单一登录](../outlook/implement-sso-in-outlook-add-in.md)

## <a name="distributing-sso-enabled-add-ins-in-microsoft-appsource"></a>在 Microsoft AppSource 中分发支持 SSO 的加载项

当 Microsoft 365 管理员从[AppSource](https://appsource.microsoft.com)获取加载项时，管理员可以通过集中部署重新分发[](../publish/centralized-deployment.md)它，并授予管理员对加载项的访问权限以访问 Microsoft Graph 范围。 但是，最终用户也可以直接从 AppSource 获取外接程序，在这种情况下，用户必须同意加载项。 这可以创建一个潜在的性能问题，我们提供了一个解决方案。

如果你的代码在调用中传递选项，例如，如果 Azure AD 向 Office 报告尚未向加载项授予同意，Office 可以提示用户同意 `allowConsentPrompt` `getAccessToken` `OfficeRuntime.auth.getAccessToken( { allowConsentPrompt: true } );` 。 但是，出于安全原因，Office 只能提示用户同意 Azure AD `profile` 作用域。 *Office 无法提示同意任何 Microsoft Graph 范围，* 甚至无法 `User.Read` 提示。 这意味着，如果用户同意提示，Office 将返回启动令牌。 但是，尝试将启动令牌交换为 Microsoft Graph 访问令牌将失败，出现错误 AADSTS65001，这意味着尚未授予 (Microsoft Graph 作用域) 同意。

你的代码可以并且应该通过返回到备用身份验证系统来处理此错误，这将提示用户同意 Microsoft Graph 作用域。  (有关代码示例，请参阅"创建使用单一登录的 [Node.js Office](create-sso-office-add-ins-nodejs.md) 外接程序"和"创建使用单一登录的 [ASP.NET Office](create-sso-office-add-ins-aspnet.md) 外接程序及其链接到的示例。) 但整个过程需要多次往返 Azure AD。 可以通过在调用中包括选项来避免这种性能损失; `forMSGraphAccess` `getAccessToken` 例如， `OfficeRuntime.auth.getAccessToken( { forMSGraphAccess: true } )` 。  这向 Office 发出加载项需要 Microsoft Graph 作用域的信号。 Office 将要求 Azure AD 验证是否向加载项授予了对 Microsoft Graph 范围的同意。 如果有，将返回启动令牌。 如果尚未调用，则调用将返回错误 `getAccessToken` 13012。 代码可以通过立即返回到备用身份验证系统来处理此错误，而无需尝试与 Azure AD 交换令牌。

最佳做法是，始终传递加载项在 AppSource 中分发并 `forMSGraphAccess` `getAccessToken` 需要 Microsoft Graph 范围时。

> [!TIP]
> 如果您开发使用 SSO 的 Outlook 外接程序并旁加载它进行测试，则即使已授予管理员同意，Office 也会始终返回错误 13012。 `forMSGraphAccess` `getAccessToken` 因此，在开发 Outlook 外接程序时，应 `forMSGraphAccess` 注释掉此选项。  在部署用于生产时，请务必取消对该选项的注释。 只有在 Outlook 中旁加载时，才发生假名 13012。
